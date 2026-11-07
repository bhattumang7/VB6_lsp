//! Finding the file of a registered type library.
//!
//! The registry lists a library under `TypeLib\{guid}\<major>.<minor>` (both in
//! hexadecimal), then a language id, then the platform (`win32`, `win64`) whose
//! default value is the path of the file that holds it. The selection rules are
//! plain functions so they can be exercised without a registry.

use std::path::PathBuf;

/// A registered version of a library, as the registry names it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RegisteredVersion {
    pub major: u16,
    pub minor: u16,
}

/// Parse a registry version key (`2.8`, `6.1`, `a.0`): two hexadecimal numbers.
pub fn parse_version_key(key: &str) -> Option<RegisteredVersion> {
    let (major, minor) = key.split_once('.')?;
    Some(RegisteredVersion {
        major: u16::from_str_radix(major, 16).ok()?,
        minor: u16::from_str_radix(minor, 16).ok()?,
    })
}

/// The registry key name of a version: the inverse of [`parse_version_key`].
pub fn version_key(version: RegisteredVersion) -> String {
    format!("{:x}.{:x}", version.major, version.minor)
}

/// The registered version a request for `major.minor` is served by: the same
/// major version with the smallest minor version that is not older than the one
/// asked for (an exact match wins).
pub fn select_version(registered: &[RegisteredVersion], major: u16, minor: u16) -> Option<RegisteredVersion> {
    registered
        .iter()
        .copied()
        .filter(|v| v.major == major && v.minor >= minor)
        .min_by_key(|v| v.minor)
}

/// The language-id keys to try for a request, best first: the language asked
/// for, then the neutral language, then any other that is registered.
pub fn language_candidates(requested: u32, registered: &[String]) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    let mut push = |key: String| {
        if registered.iter().any(|r| r.eq_ignore_ascii_case(&key)) && !out.iter().any(|o| o.eq_ignore_ascii_case(&key)) {
            out.push(key);
        }
    };
    push(format!("{requested:x}"));
    push("0".to_string());
    for key in registered {
        if key.chars().all(|c| c.is_ascii_hexdigit()) {
            push(key.clone());
        }
    }
    out
}

/// Replace `%NAME%` references with the value `lookup` gives; a reference it
/// does not know is left as written.
pub fn expand_variables(text: &str, lookup: impl Fn(&str) -> Option<String>) -> String {
    let mut out = String::new();
    let mut rest = text;
    while let Some(start) = rest.find('%') {
        out.push_str(&rest[..start]);
        let after = &rest[start + 1..];
        match after.find('%') {
            Some(end) => {
                let name = &after[..end];
                match lookup(name) {
                    Some(value) if !name.is_empty() => out.push_str(&value),
                    _ => {
                        out.push('%');
                        out.push_str(name);
                        out.push('%');
                    }
                }
                rest = &after[end + 1..];
            }
            None => {
                out.push('%');
                rest = after;
            }
        }
    }
    out.push_str(rest);
    out
}

/// The file path a registry value names. The value may end in `\<n>`, which
/// selects the n-th type-library resource of a multi-library file: the path
/// without it is what exists on disk, and the whole value is what loads.
pub fn registered_path(value: &str) -> PathBuf {
    PathBuf::from(expand_variables(value.trim().trim_matches('"'), |name| std::env::var(name).ok()))
}

/// The on-disk file of a registry value, without any resource-index suffix.
pub fn file_of_registered(value: &str) -> PathBuf {
    let path = registered_path(value);
    if path.exists() {
        return path;
    }
    let text = path.to_string_lossy().into_owned();
    if let Some((file, index)) = text.rsplit_once('\\') {
        if !index.is_empty() && index.chars().all(|c| c.is_ascii_digit()) {
            return PathBuf::from(file);
        }
    }
    path
}
