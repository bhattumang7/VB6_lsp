//! Choosing a registered type library and reading the path the registry gives for it.

use std::path::PathBuf;
use vb6_lsp::typelib::locate::*;

fn v(major: u16, minor: u16) -> RegisteredVersion {
    RegisteredVersion { major, minor }
}

#[test]
fn version_keys_are_two_hexadecimal_numbers() {
    assert_eq!(parse_version_key("2.8"), Some(v(2, 8)));
    assert_eq!(parse_version_key("a.1f"), Some(v(10, 31)));
    assert_eq!(parse_version_key("6.0"), Some(v(6, 0)));
    assert_eq!(parse_version_key("6"), None);
    assert_eq!(parse_version_key("x.1"), None);
    assert_eq!(parse_version_key("1.10000"), None);
}

#[test]
fn a_version_key_round_trips() {
    assert_eq!(version_key(v(2, 8)), "2.8");
    assert_eq!(version_key(v(10, 31)), "a.1f");
    assert_eq!(parse_version_key(&version_key(v(65535, 255))), Some(v(65535, 255)));
}

#[test]
fn the_exact_version_wins_then_the_smallest_newer_minor_of_the_same_major() {
    let registered = [v(2, 0), v(2, 5), v(2, 8), v(6, 0)];
    assert_eq!(select_version(&registered, 2, 5), Some(v(2, 5)));
    assert_eq!(select_version(&registered, 2, 6), Some(v(2, 8)));
    assert_eq!(select_version(&registered, 2, 1), Some(v(2, 5)));
    assert_eq!(select_version(&registered, 2, 9), None);
    assert_eq!(select_version(&registered, 6, 0), Some(v(6, 0)));
    assert_eq!(select_version(&registered, 3, 0), None);
    assert_eq!(select_version(&[], 2, 0), None);
}

#[test]
fn the_requested_language_is_tried_before_the_neutral_one_and_any_other() {
    let registered: Vec<String> = ["409", "0", "FLAGS", "HELPDIR", "407"].iter().map(|s| s.to_string()).collect();
    assert_eq!(language_candidates(0x409, &registered), vec!["409", "0", "407"]);
    assert_eq!(language_candidates(0, &registered), vec!["0", "409", "407"]);
    assert_eq!(language_candidates(0x411, &registered), vec!["0", "409", "407"]);
    assert_eq!(language_candidates(0x409, &[]), Vec::<String>::new());
}

fn known(name: &str) -> Option<String> {
    match name {
        "windir" => Some(r"C:\Windows".to_string()),
        "EMPTY" => Some(String::new()),
        _ => None,
    }
}

#[test]
fn variables_in_a_registered_path_are_expanded() {
    assert_eq!(expand_variables(r"%windir%\system32\x.dll", known), r"C:\Windows\system32\x.dll");
    assert_eq!(expand_variables(r"%unknown%\x.dll", known), r"%unknown%\x.dll");
    assert_eq!(expand_variables(r"100%\x.dll", known), r"100%\x.dll");
    assert_eq!(expand_variables("a%EMPTY%b", known), "ab");
    assert_eq!(expand_variables("%%", known), "%%");
    assert_eq!(expand_variables("%windir%%windir%", known), r"C:\WindowsC:\Windows");
    assert_eq!(expand_variables("plain", known), "plain");
}

#[test]
fn a_registered_value_is_trimmed_and_unquoted() {
    assert_eq!(registered_path("  \"C:\\Lib\\x.tlb\"  "), PathBuf::from(r"C:\Lib\x.tlb"));
}

#[test]
fn a_resource_index_is_dropped_to_find_the_file_on_disk() {
    let dir = std::env::temp_dir().join(format!("vb6lsp-locate-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let file = dir.join("multi.dll");
    std::fs::write(&file, b"x").unwrap();
    let with_index = format!("{}\\2", file.display());
    assert_eq!(file_of_registered(&with_index), file);
    assert_eq!(file_of_registered(&file.display().to_string()), file);
    let absent = dir.join("absent.dll");
    assert_eq!(file_of_registered(&absent.display().to_string()), absent);
    std::fs::remove_dir_all(&dir).unwrap();
}
