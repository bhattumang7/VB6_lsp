//! Shared by the project tests.

#![allow(dead_code)]

/// A scratch directory holding a project's files, removed when dropped.
pub struct TempDir {
    root: std::path::PathBuf,
}

impl TempDir {
    pub fn new(label: &str) -> Self {
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("vb6proj-{label}-{}-{nanos}", std::process::id()));
        std::fs::create_dir_all(&root).unwrap();
        // Resolve the spelling the file system reports (short names, symlinks).
        let root = std::fs::canonicalize(&root).map(strip_verbatim).unwrap_or(root);
        TempDir { root }
    }

    pub fn path(&self, rel: &str) -> std::path::PathBuf {
        self.root.join(rel)
    }

    /// Write `content` (Windows-1252 text; ASCII in these tests) to `rel`, creating folders.
    pub fn write(&self, rel: &str, content: &str) -> std::path::PathBuf {
        let path = self.path(rel);
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, content).unwrap();
        path
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.root);
    }
}

/// `\?\C:\x` becomes `C:\x`.
fn strip_verbatim(path: std::path::PathBuf) -> std::path::PathBuf {
    match path.to_string_lossy().strip_prefix(r"\?\") {
        Some(rest) => std::path::PathBuf::from(rest),
        None => path,
    }
}

/// The message for a library that is neither registered nor at `tried`, as this platform words it.
pub fn missing_library_message(what: &str, tried: &str) -> String {
    if cfg!(windows) {
        format!("Type library {what} is not registered and was not found at {tried}. Its types are not available to the analysis")
    } else {
        format!("Type library {what} cannot be read on this platform. Its types are not available to the analysis")
    }
}
