//! Runs the installed VB6 compiler headlessly on single-module projects and
//! turns its build log into a [`Verdict`]. Verdicts are cached on disk, keyed by
//! the exact source bytes and the compiler build.

use std::fs;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};

/// Starting a process without a console window of its own.
trait NoWindow {
    fn creation_flags_no_window(&mut self) -> &mut Self;
}

impl NoWindow for Command {
    #[cfg(windows)]
    fn creation_flags_no_window(&mut self) -> &mut Self {
        use std::os::windows::process::CommandExt;
        const CREATE_NO_WINDOW: u32 = 0x0800_0000;
        self.creation_flags(CREATE_NO_WINDOW)
    }

    #[cfg(not(windows))]
    fn creation_flags_no_window(&mut self) -> &mut Self {
        self
    }
}

/// Kind of module a case is compiled as.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ModuleKind {
    Bas,
    Cls,
}

impl ModuleKind {
    fn extension(self) -> &'static str {
        match self {
            ModuleKind::Bas => "bas",
            ModuleKind::Cls => "cls",
        }
    }
}

/// One input for the compiler: an id for reporting, the module kind and the
/// exact file bytes.
#[derive(Clone, Debug)]
pub struct Case {
    pub id: String,
    pub kind: ModuleKind,
    pub source: Vec<u8>,
}

impl Case {
    /// Builds a case from text, normalising every line break to CRLF.
    pub fn from_text(id: impl Into<String>, kind: ModuleKind, text: &str) -> Case {
        Case { id: id.into(), kind, source: normalize_crlf(text.as_bytes()) }
    }
}

/// Rewrites every line break (`LF`, `CR` or `CRLF`) as `CRLF`. A case that holds
/// bytes of its own is written as it is: the compiler reads every kind of line break.
pub fn normalize_crlf(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(bytes.len() + 16);
    let mut i = 0;
    while i < bytes.len() {
        match bytes[i] {
            b'\r' if bytes.get(i + 1) == Some(&b'\n') => {
                out.extend_from_slice(b"\r\n");
                i += 2;
            }
            b'\r' | b'\n' => {
                out.extend_from_slice(b"\r\n");
                i += 1;
            }
            other => {
                out.push(other);
                i += 1;
            }
        }
    }
    out
}

/// The compiler's answer for one case.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Verdict {
    Accepted,
    /// `Compile Error in File '<file>', Line <line> : <message>`. The line is
    /// exactly as the compiler reports it, which counts from 0.
    Rejected { file: String, line: Option<u32>, message: String },
    /// The build failed with a message that names no file or line (project level,
    /// such as an ambiguous name between modules).
    ProjectError { message: String },
    /// The log held neither a success nor a recognised error line; the raw log
    /// and exit code are kept so nothing is silently dropped.
    Unrecognised { exit_code: Option<i32>, log: String },
    /// The compiler did not finish within the time limit.
    TimedOut,
}

impl Verdict {
    fn to_json(&self) -> serde_json::Value {
        use serde_json::json;
        match self {
            Verdict::Accepted => json!({"v": "accepted"}),
            Verdict::Rejected { file, line, message } => {
                json!({"v": "rejected", "file": file, "line": line, "message": message})
            }
            Verdict::ProjectError { message } => json!({"v": "project_error", "message": message}),
            Verdict::Unrecognised { exit_code, log } => {
                json!({"v": "unrecognised", "exit": exit_code, "log": log})
            }
            Verdict::TimedOut => json!({"v": "timed_out"}),
        }
    }

    fn from_json(value: &serde_json::Value) -> Option<Verdict> {
        Some(match value.get("v")?.as_str()? {
            "accepted" => Verdict::Accepted,
            "rejected" => Verdict::Rejected {
                file: value.get("file")?.as_str()?.to_string(),
                line: value.get("line")?.as_u64().map(|n| n as u32),
                message: value.get("message")?.as_str()?.to_string(),
            },
            "project_error" => Verdict::ProjectError { message: value.get("message")?.as_str()?.to_string() },
            "unrecognised" => Verdict::Unrecognised {
                exit_code: value.get("exit")?.as_i64().map(|n| n as i32),
                log: value.get("log")?.as_str()?.to_string(),
            },
            "timed_out" => Verdict::TimedOut,
            _ => return None,
        })
    }
}

/// Interprets the build log written by `/make ... /out`.
pub fn parse_log(log: &str, exit_code: Option<i32>) -> Verdict {
    let mut rejected = None;
    let mut succeeded = false;
    let mut failed = false;
    let mut other = Vec::new();
    for line in log.lines() {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        if let Some(rest) = line.strip_prefix("Compile Error in File '") {
            // `<file>', Line <n> : <message>`  (the line part can be absent)
            if let Some((file, tail)) = rest.split_once("', ") {
                let (line_no, message) = match tail.strip_prefix("Line ") {
                    Some(t) => match t.split_once(" : ") {
                        Some((n, m)) => (n.trim().parse().ok(), m.to_string()),
                        None => (None, t.to_string()),
                    },
                    None => (None, tail.trim_start_matches(": ").to_string()),
                };
                rejected.get_or_insert(Verdict::Rejected {
                    file: file.to_string(),
                    line: line_no,
                    message,
                });
            }
        } else if line.starts_with("Build of '") && line.ends_with("succeeded.") {
            succeeded = true;
        } else if line.starts_with("Build of '") && line.ends_with("failed.") {
            failed = true;
        } else {
            other.push(line.to_string());
        }
    }
    match (rejected, succeeded, exit_code) {
        (Some(v), _, Some(1)) => v,
        (None, true, Some(0)) => Verdict::Accepted,
        (None, false, Some(1)) if failed && !other.is_empty() => {
            Verdict::ProjectError { message: other.join(" | ") }
        }
        _ => Verdict::Unrecognised { exit_code, log: log.to_string() },
    }
}

/// Startup module compiled alongside every case.
const HARNESS_MAIN: &str = "Attribute VB_Name = \"HarnessMain\"\r\nSub Main()\r\nEnd Sub\r\n";

/// Settings for the oracle runner.
pub struct Oracle {
    pub vb6_exe: PathBuf,
    pub work_root: PathBuf,
    pub cache_dir: PathBuf,
    pub workers: usize,
    pub timeout: Duration,
    /// Mixed into every cache key so a different compiler build never reuses
    /// another build's verdicts.
    build_salt: String,
}

fn fnv1a64(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf29ce484222325u64;
    for b in bytes {
        hash ^= u64::from(*b);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

impl Oracle {
    pub fn new(vb6_exe: PathBuf, work_root: PathBuf, cache_dir: PathBuf) -> std::io::Result<Oracle> {
        let meta = fs::metadata(&vb6_exe)?;
        let modified = meta
            .modified()?
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs())
            .unwrap_or(0);
        let build_salt = format!("{}:{}:{}:{}", vb6_exe.display(), meta.len(), modified, HARNESS_MAIN);
        fs::create_dir_all(&work_root)?;
        fs::create_dir_all(&cache_dir)?;
        Ok(Oracle {
            vb6_exe,
            work_root,
            cache_dir,
            // `ORACLE_WORKERS` sets how many compiler instances run at once; the default
            // is modest because every instance starts and stops many times a second.
            workers: std::env::var("ORACLE_WORKERS")
                .ok()
                .and_then(|v| v.parse().ok())
                .filter(|n| *n > 0)
                .unwrap_or(3),
            timeout: Duration::from_secs(90),
            build_salt,
        })
    }

    /// Cache entry path for a case plus the exact key bytes stored inside it.
    fn cache_entry(&self, case: &Case) -> (PathBuf, Vec<u8>) {
        let mut key = Vec::new();
        key.extend_from_slice(self.build_salt.as_bytes());
        key.push(0);
        key.extend_from_slice(case.kind.extension().as_bytes());
        key.push(0);
        key.extend_from_slice(&case.source);
        let name = format!("{:016x}-{}.json", fnv1a64(&key), key.len());
        (self.cache_dir.join(name), key)
    }

    fn cache_get(&self, case: &Case) -> Option<Verdict> {
        let (path, key) = self.cache_entry(case);
        let entry: serde_json::Value = serde_json::from_slice(&fs::read(path).ok()?).ok()?;
        // The hash only names the file; equality of the stored key decides a hit.
        let stored = entry.get("key")?.as_array()?;
        if stored.len() != key.len()
            || !stored.iter().zip(&key).all(|(a, b)| a.as_u64() == Some(u64::from(*b)))
        {
            return None;
        }
        Verdict::from_json(entry.get("verdict")?)
    }

    fn cache_put(&self, case: &Case, verdict: &Verdict) {
        // Timeouts and unrecognised logs are not facts about the source, so they are not cached.
        if matches!(verdict, Verdict::TimedOut | Verdict::Unrecognised { .. }) {
            return;
        }
        let (path, key) = self.cache_entry(case);
        let entry = serde_json::json!({"key": key, "verdict": verdict.to_json()});
        let _ = fs::write(path, serde_json::to_vec(&entry).unwrap_or_default());
    }

    /// Compiles one case in its own directory.
    fn compile(&self, slot: usize, case: &Case) -> Verdict {
        match self.compile_once(slot, case, false) {
            // The module defines its own `Main`: build again with that as the only startup.
            Verdict::ProjectError { message } if message == "Ambiguous name detected: Main" => {
                self.compile_once(slot, case, true)
            }
            other => other,
        }
    }

    fn compile_once(&self, slot: usize, case: &Case, own_main: bool) -> Verdict {
        let dir = self.work_root.join(format!("w{slot}"));
        let _ = fs::remove_dir_all(&dir);
        if fs::create_dir_all(&dir).is_err() {
            return Verdict::Unrecognised { exit_code: None, log: "work directory unavailable".into() };
        }
        let module = format!("m.{}", case.kind.extension());
        let module_line = match case.kind {
            ModuleKind::Bas => format!("Module=Module1; {module}"),
            ModuleKind::Cls => format!("Class=Class1; {module}"),
        };
        // The harness supplies the startup module, so a case need not define `Main`.
        let harness_line = if own_main { "" } else { "Module=HarnessMain; h.bas\r\n" };
        let project = format!(
            "Type=Exe\r\n{module_line}\r\n{harness_line}Startup=\"Sub Main\"\r\nExeName32=\"p.exe\"\r\n"
        );
        if (!own_main && fs::write(dir.join("h.bas"), HARNESS_MAIN).is_err())
            || fs::write(dir.join(&module), &case.source).is_err()
            || fs::write(dir.join("p.vbp"), project).is_err()
        {
            return Verdict::Unrecognised { exit_code: None, log: "case files unwritable".into() };
        }
        let log_path = dir.join("p.log");
        let mut child = match Command::new(&self.vb6_exe)
            .arg("/make")
            .arg(dir.join("p.vbp"))
            .arg("/out")
            .arg(&log_path)
            .current_dir(&dir)
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .creation_flags_no_window()
            .spawn()
        {
            Ok(c) => c,
            Err(e) => {
                return Verdict::Unrecognised { exit_code: None, log: format!("spawn failed: {e}") }
            }
        };
        let started = Instant::now();
        let status = loop {
            match child.try_wait() {
                Ok(Some(s)) => break s,
                Ok(None) if started.elapsed() > self.timeout => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Verdict::TimedOut;
                }
                Ok(None) => std::thread::sleep(Duration::from_millis(20)),
                Err(e) => {
                    return Verdict::Unrecognised { exit_code: None, log: format!("wait failed: {e}") }
                }
            }
        };
        let log = fs::read(&log_path)
            .map(|b| String::from_utf8_lossy(&b).into_owned())
            .unwrap_or_default();
        parse_log(&log, status.code())
    }

    /// Verdicts for every case, in input order. Cached verdicts are reused; the
    /// rest are compiled by `workers` parallel compiler instances.
    pub fn run(&self, cases: &[Case]) -> Vec<Verdict> {
        let results: Mutex<Vec<Option<Verdict>>> = Mutex::new(vec![None; cases.len()]);
        let mut pending = Vec::new();
        for (i, case) in cases.iter().enumerate() {
            match self.cache_get(case) {
                Some(v) => results.lock().unwrap()[i] = Some(v),
                None => pending.push(i),
            }
        }
        let next = AtomicUsize::new(0);
        std::thread::scope(|scope| {
            for slot in 0..self.workers.max(1) {
                let (results, pending, next) = (&results, &pending, &next);
                scope.spawn(move || loop {
                    let n = next.fetch_add(1, Ordering::SeqCst);
                    let Some(&i) = pending.get(n) else { break };
                    let verdict = self.compile(slot, &cases[i]);
                    self.cache_put(&cases[i], &verdict);
                    results.lock().unwrap()[i] = Some(verdict);
                });
            }
        });
        results.into_inner().unwrap().into_iter().map(|v| v.expect("every case has a verdict")).collect()
    }
}

/// Default location of the installed compiler.
pub fn default_vb6_exe() -> PathBuf {
    PathBuf::from(r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE")
}

/// Directory helper used by the CLI and tests.
pub fn ensure_dir(path: &Path) -> std::io::Result<()> {
    fs::create_dir_all(path)
}
