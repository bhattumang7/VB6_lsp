//! Encoding detection and handling for VB6 files
//!
//! VB6 files can be in either:
//! - UTF-8 (modern editors, converted files)
//! - Windows-1252 / CP1252 (original VB6 IDE default)
//!
//! This module provides utilities to:
//! 1. Detect the encoding of a file
//! 2. Read files with proper encoding handling
//! 3. Preserve the original encoding for future writes

use encoding_rs::WINDOWS_1252;
use std::fs;
use std::io;
use std::path::Path;
use tracing::{debug, warn};

/// Represents the detected or known encoding of a VB6 file
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Encoding {
    /// UTF-8 encoding (modern, web-compatible)
    Utf8,
    /// Windows-1252 / CP1252 (VB6 IDE default)
    Windows1252,
}

impl Encoding {
    /// Get a human-readable name for this encoding
    pub fn name(&self) -> &'static str {
        match self {
            Encoding::Utf8 => "UTF-8",
            Encoding::Windows1252 => "Windows-1252",
        }
    }
}

/// Represents the content of a VB6 file along with its detected encoding
#[derive(Debug, Clone)]
pub struct VB6FileContent {
    /// The text content of the file
    pub text: String,
    /// The detected encoding
    pub encoding: Encoding,
    /// Whether there were any encoding errors during decoding
    pub had_errors: bool,
}

/// Utility for reading VB6 files with encoding detection
pub struct VB6FileReader;

impl VB6FileReader {
    /// Read a VB6 file from disk with automatic encoding detection
    ///
    /// This function:
    /// 1. Reads the file as bytes
    /// 2. Tries UTF-8 first (no conversion needed if valid)
    /// 3. Falls back to Windows-1252 if UTF-8 validation fails
    /// 4. Returns the decoded text along with the detected encoding
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use std::path::Path;
    /// use vb6_lsp::utils::VB6FileReader;
    ///
    /// let content = VB6FileReader::read_file(Path::new("Module1.bas")).unwrap();
    /// println!("Read {} with encoding: {}", content.text.len(), content.encoding.name());
    /// ```
    pub fn read_file(path: &Path) -> io::Result<VB6FileContent> {
        debug!("Reading VB6 file: {}", path.display());

        // Read the raw bytes
        let bytes = fs::read(path)?;

        // Detect and decode
        let content = Self::detect_and_decode(&bytes, path);

        debug!(
            "File {} decoded as {} ({} bytes, {} errors)",
            path.display(),
            content.encoding.name(),
            content.text.len(),
            if content.had_errors { "had" } else { "no" }
        );

        Ok(content)
    }

    /// Detect encoding and decode bytes to a string
    ///
    /// Priority order:
    /// 1. Try UTF-8 (lossless, no BOM required)
    /// 2. Fall back to Windows-1252 (VB6 default)
    ///
    /// # Arguments
    ///
    /// * `bytes` - Raw file bytes
    /// * `path` - File path (for logging only)
    pub fn detect_and_decode(bytes: &[u8], path: &Path) -> VB6FileContent {
        // Check for UTF-8 BOM (EF BB BF)
        let has_utf8_bom = bytes.starts_with(&[0xEF, 0xBB, 0xBF]);

        if has_utf8_bom {
            debug!("File {} has UTF-8 BOM", path.display());
            let text = String::from_utf8_lossy(&bytes[3..]).into_owned();
            return VB6FileContent {
                text,
                encoding: Encoding::Utf8,
                had_errors: false,
            };
        }

        // Try UTF-8 without BOM
        match String::from_utf8(bytes.to_vec()) {
            Ok(text) => {
                // Successfully decoded as UTF-8
                debug!("File {} is valid UTF-8", path.display());
                VB6FileContent {
                    text,
                    encoding: Encoding::Utf8,
                    had_errors: false,
                }
            }
            Err(_) => {
                // Not valid UTF-8, try Windows-1252
                debug!(
                    "File {} is not UTF-8, attempting Windows-1252 decode",
                    path.display()
                );
                Self::decode_windows1252(bytes, path)
            }
        }
    }

    /// Decode bytes as Windows-1252
    ///
    /// Windows-1252 (CP1252) is the default encoding used by the VB6 IDE.
    /// It's a superset of ISO-8859-1 (Latin-1) with additional characters
    /// in the 0x80-0x9F range.
    fn decode_windows1252(bytes: &[u8], path: &Path) -> VB6FileContent {
        let (decoded, _, had_errors) = WINDOWS_1252.decode(bytes);

        if had_errors {
            warn!(
                "File {} had decoding errors when reading as Windows-1252",
                path.display()
            );
        }

        VB6FileContent {
            text: decoded.into_owned(),
            encoding: Encoding::Windows1252,
            had_errors,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_utf8_detection() {
        let text = "Option Explicit\r\n";
        let bytes = text.as_bytes();

        let content = VB6FileReader::detect_and_decode(bytes, Path::new("test.bas"));

        assert_eq!(content.encoding, Encoding::Utf8);
        assert_eq!(content.text, text);
        assert!(!content.had_errors);
    }

    #[test]
    fn test_utf8_bom_detection() {
        let text = "Option Explicit\r\n";
        let mut bytes = vec![0xEF, 0xBB, 0xBF]; // UTF-8 BOM
        bytes.extend_from_slice(text.as_bytes());

        let content = VB6FileReader::detect_and_decode(&bytes, Path::new("test.bas"));

        assert_eq!(content.encoding, Encoding::Utf8);
        assert_eq!(content.text, text);
        assert!(!content.had_errors);
    }

    #[test]
    fn test_windows1252_detection() {
        // Create a byte sequence with Windows-1252 specific character
        // 0x93 is a left double quotation mark in Windows-1252
        let bytes = vec![
            0x4F, 0x70, 0x74, 0x69, 0x6F, 0x6E, 0x20, // "Option "
            0x93, // Left double quote (Windows-1252)
            0x45, 0x78, 0x70, 0x6C, 0x69, 0x63, 0x69, 0x74, // "Explicit"
            0x94, // Right double quote (Windows-1252)
        ];

        let content = VB6FileReader::detect_and_decode(&bytes, Path::new("test.bas"));

        assert_eq!(content.encoding, Encoding::Windows1252);
        assert!(!content.text.is_empty());
    }
}
