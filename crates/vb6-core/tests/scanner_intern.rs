//! Tests for keyword interning and the scanner context.

use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::frontend::token::Kw;

#[test]
fn intern_keywords_populates_context() {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();

    // Check well-known keywords via case-insensitive name lookup.
    assert!(c.lookup("sub").is_some());
    assert!(c.lookup("SUB").is_some()); // case-insensitive
    assert!(c.lookup("If").is_some());

    // Verify the token id for "dim" matches Kw::Dim.
    if let Some(sym) = c.lookup("dim") {
        assert_eq!(sym.token as u16, Kw::Dim.token_id());
    } else {
        panic!("'dim' should be a keyword");
    }
}

#[test]
fn non_keyword_lookup_returns_none() {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    assert!(c.lookup("not_a_keyword").is_none());
}

#[test]
fn for_keyword_is_present() {
    // The keyword 'for' is present in the table.
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    assert!(c.lookup("For").is_some());
}
