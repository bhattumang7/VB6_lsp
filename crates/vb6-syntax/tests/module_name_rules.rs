//! The pure rules for the name an `Attribute VB_Name` line gives a module.

use vb6_syntax::frontend::module_name::{is_legal_standard_module_name, unescape, MAX_STANDARD_MODULE_NAME_LEN};

#[test]
fn unescape_applies_each_escape() {
    assert_eq!(unescape("plain"), "plain");
    assert_eq!(unescape("a\\nb"), "a\nb");
    assert_eq!(unescape("a\\rb"), "a\rb");
    assert_eq!(unescape("a\\\\b"), "a\\b");
    assert_eq!(unescape("a\\qb"), "aqb");
    assert_eq!(unescape("a\\"), "a");
    assert_eq!(unescape("\\"), "");
    assert_eq!(unescape(""), "");
    assert_eq!(unescape("\\\\n"), "\\n");
}

#[test]
fn legal_standard_module_names() {
    for name in ["a", "Z", "a1", "a_", "A_1b", "\u{e4}", "a\u{e9}"] {
        assert!(is_legal_standard_module_name(name), "{name:?}");
    }
    for name in ["", "1a", "_a", "a b", "a.b", "a-b", "a$", "a\n"] {
        assert!(!is_legal_standard_module_name(name), "{name:?}");
    }
}

#[test]
fn the_longest_legal_standard_module_name() {
    assert_eq!(MAX_STANDARD_MODULE_NAME_LEN, 127);
    assert!(is_legal_standard_module_name(&"a".repeat(127)));
    assert!(!is_legal_standard_module_name(&"a".repeat(128)));
    assert!(is_legal_standard_module_name(&"\u{e4}".repeat(127)));
    assert!(!is_legal_standard_module_name(&"\u{e4}".repeat(128)));
}
