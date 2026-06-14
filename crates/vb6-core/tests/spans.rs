//! Tests for source spans and location tracking.

use vb6_core::frontend::scanner::{Scanner, ScannerContext};

#[test]
fn span_tracking_basic() {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    let src = b"Dim x";
    let mut s = Scanner::new(&mut c, src);

    let t1 = s.next_token(); // Dim
    assert_eq!(t1.span.start, 0);
    assert_eq!(t1.span.start + t1.span.len, 3);

    let t2 = s.next_token(); // x
    assert_eq!(t2.span.start, 4);
    assert_eq!(t2.span.start + t2.span.len, 5);
}
