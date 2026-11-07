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
