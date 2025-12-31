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

use encoding_rs::{Encoding as EncodingRs, UTF_8, WINDOWS_1252};
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
    /// Unknown or mixed encoding
    Unknown,
}

impl Encoding {
    /// Get the encoding_rs::Encoding for this type
    pub fn as_encoding_rs(&self) -> &'static EncodingRs {
        match self {
            Encoding::Utf8 => UTF_8,
            Encoding::Windows1252 => WINDOWS_1252,
            Encoding::Unknown => WINDOWS_1252, // Default fallback
        }
    }

    /// Get a human-readable name for this encoding
    pub fn name(&self) -> &'static str {
        match self {
            Encoding::Utf8 => "UTF-8",
            Encoding::Windows1252 => "Windows-1252",
            Encoding::Unknown => "Unknown (fallback to Windows-1252)",
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

    /// Read a VB6 file and return just the text content
    ///
    /// This is a convenience method that discards encoding information.
    /// Use `read_file()` if you need to know or preserve the encoding.
    pub fn read_to_string(path: &Path) -> io::Result<String> {
        Ok(Self::read_file(path)?.text)
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

    /// Encode a string back to bytes using the specified encoding
    ///
    /// Use this when writing VB6 files to preserve their original encoding.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use vb6_lsp::utils::{VB6FileReader, Encoding};
    ///
    /// let text = "Option Explicit\r\n";
    /// let bytes = VB6FileReader::encode_string(text, Encoding::Windows1252);
    /// // Write bytes to file...
    /// ```
    pub fn encode_string(text: &str, encoding: Encoding) -> Vec<u8> {
        match encoding {
            Encoding::Utf8 => text.as_bytes().to_vec(),
            Encoding::Windows1252 | Encoding::Unknown => {
                let (encoded, _, _) = WINDOWS_1252.encode(text);
                encoded.into_owned()
            }
        }
    }

    /// Check if bytes are likely to be Windows-1252 encoded
    ///
    /// This is a heuristic check that looks for:
    /// - Bytes in the 0x80-0x9F range (Windows-1252 specific)
    /// - Invalid UTF-8 sequences
    ///
    /// Returns `true` if the bytes are likely Windows-1252.
    pub fn is_likely_windows1252(bytes: &[u8]) -> bool {
        // If it's valid UTF-8, it's probably UTF-8
        if String::from_utf8(bytes.to_vec()).is_ok() {
            return false;
        }

        // Look for Windows-1252 specific characters (0x80-0x9F)
        // These are control characters in ISO-8859-1 but printable in Windows-1252
        let has_win1252_chars = bytes.iter().any(|&b| matches!(b, 0x80..=0x9F));

        has_win1252_chars
    }

    /// Detect encoding without decoding the entire file
    ///
    /// This is useful for large files where you only need to know
    /// the encoding without actually reading the content.
    pub fn detect_encoding(bytes: &[u8]) -> Encoding {
        // Check for UTF-8 BOM
        if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
            return Encoding::Utf8;
        }

        // Try UTF-8 validation
        if String::from_utf8(bytes.to_vec()).is_ok() {
            return Encoding::Utf8;
        }

        // Check for Windows-1252 indicators
        if Self::is_likely_windows1252(bytes) {
            return Encoding::Windows1252;
        }

        Encoding::Unknown
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

    #[test]
    fn test_encode_utf8() {
        let text = "Option Explicit";
        let bytes = VB6FileReader::encode_string(text, Encoding::Utf8);

        assert_eq!(bytes, text.as_bytes());
    }

    #[test]
    fn test_encode_windows1252() {
        let text = "Option Explicit";
        let bytes = VB6FileReader::encode_string(text, Encoding::Windows1252);

        // Should be able to decode it back
        let (decoded, _, had_errors) = WINDOWS_1252.decode(&bytes);
        assert_eq!(decoded, text);
        assert!(!had_errors);
    }

    #[test]
    fn test_is_likely_windows1252() {
        // UTF-8 text
        let utf8_bytes = "Option Explicit".as_bytes();
        assert!(!VB6FileReader::is_likely_windows1252(utf8_bytes));

        // Windows-1252 with special char
        let win1252_bytes = vec![0x4F, 0x70, 0x74, 0x93, 0x45];
        assert!(VB6FileReader::is_likely_windows1252(&win1252_bytes));
    }
}
