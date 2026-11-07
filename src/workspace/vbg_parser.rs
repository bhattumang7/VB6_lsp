//! VBG project-group file parser.
//!
//! A `.vbg` file is a line-oriented text file: a `VBGROUP 5.0` header, an optional
//! `StartupProject=` line and one `Project=` line per project in the group. Paths
//! are relative to the group file unless absolute.

use std::path::{Path, PathBuf};

use crate::utils::VB6FileReader;

/// A project listed in a group.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GroupProject {
    /// The path as written in the group file.
    pub relative_path: PathBuf,
    /// The path resolved against the group file's directory.
    pub absolute_path: PathBuf,
    /// 1-based line of the `Project=` entry.
    pub line: usize,
}

/// A parsed `.vbg` file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbgFile {
    /// The version in the header (`5.0`).
    pub version: String,
    /// The projects of the group, in file order.
    pub projects: Vec<GroupProject>,
    /// The `StartupProject=` entry, resolved; `None` when the file has none.
    pub startup: Option<PathBuf>,
    /// 1-based line of the `StartupProject=` entry.
    pub startup_line: Option<usize>,
}

/// Why a group file could not be read.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbgParseError {
    pub message: String,
    pub line: Option<usize>,
}

impl std::fmt::Display for VbgParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.line {
            Some(line) => write!(f, "VBG parse error at line {}: {}", line, self.message),
            None => write!(f, "VBG parse error: {}", self.message),
        }
    }
}

impl std::error::Error for VbgParseError {}

impl VbgFile {
    /// Read and parse a group file from disk.
    pub fn parse(vbg_path: &Path) -> Result<Self, VbgParseError> {
        let file = VB6FileReader::read_file(vbg_path).map_err(|e| VbgParseError {
            message: format!("Failed to read file: {}", e),
            line: None,
        })?;
        Self::parse_content(vbg_path, &file.text)
    }

    /// Parse group-file text. `vbg_path` locates the file so relative project
    /// paths can be resolved.
    pub fn parse_content(vbg_path: &Path, content: &str) -> Result<Self, VbgParseError> {
        let dir = vbg_path.parent().unwrap_or(Path::new("."));
        let resolve = |raw: &str| {
            let path = PathBuf::from(raw);
            if path.is_absolute() {
                path
            } else {
                dir.join(path)
            }
        };

        let mut lines = content.lines().enumerate().map(|(i, l)| (i + 1, l.trim()));
        let version = loop {
            match lines.next() {
                None => {
                    return Err(VbgParseError {
                        message: "missing `VBGROUP` header".to_string(),
                        line: None,
                    })
                }
                Some((_, "")) => continue,
                Some((line, text)) => {
                    let mut words = text.split_whitespace();
                    let is_header = words.next().is_some_and(|w| w.eq_ignore_ascii_case("VBGROUP"));
                    if !is_header {
                        return Err(VbgParseError {
                            message: "the first line must be the `VBGROUP` header".to_string(),
                            line: Some(line),
                        });
                    }
                    break words.next().unwrap_or("").to_string();
                }
            }
        };

        let mut projects = Vec::new();
        let mut startup = None;
        let mut startup_line = None;
        for (line, text) in lines {
            let Some((key, value)) = text.split_once('=') else { continue };
            let value = value.trim().trim_matches('"');
            if value.is_empty() {
                continue;
            }
            match key.trim().to_ascii_lowercase().as_str() {
                "project" => projects.push(GroupProject {
                    relative_path: PathBuf::from(value),
                    absolute_path: resolve(value),
                    line,
                }),
                "startupproject" => {
                    startup = Some(resolve(value));
                    startup_line = Some(line);
                }
                _ => {}
            }
        }

        Ok(VbgFile { version, projects, startup, startup_line })
    }
}
