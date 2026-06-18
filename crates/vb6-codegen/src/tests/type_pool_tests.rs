use super::*;

#[test]
fn intern_assigns_first_seen_order_indices() {
    let mut p = TypePool::new();
    assert_eq!(p.intern(0x1234), 0);
    assert_eq!(p.intern(0x5678), 1);
    assert_eq!(p.intern(0x9abc), 2);
    assert_eq!(p.len(), 3);
}

#[test]
fn intern_deduplicates_by_type_value() {
    let mut p = TypePool::new();
    assert_eq!(p.intern(0xaa), 0);
    assert_eq!(p.intern(0xbb), 1);
    assert_eq!(p.intern(0xaa), 0); // repeat → same index
    assert_eq!(p.intern(0xbb), 1);
    assert_eq!(p.intern(0xcc), 2); // new value → next index
    assert_eq!(p.len(), 3);
}

#[test]
fn intern_handles_zero_value() {
    let mut p = TypePool::new();
    assert_eq!(p.intern(0), 0);
    assert_eq!(p.intern(0), 0);
    assert_eq!(p.intern(1), 1);
    assert_eq!(p.len(), 2);
}

#[test]
fn extract_type_value2_returns_low16_index() {
    let mut p = TypePool::new();
    assert_eq!(p.extract_type_value2(0xdead_beef), 0);
    assert_eq!(p.extract_type_value2(0xdead_beef), 0);
    assert_eq!(p.extract_type_value2(0xcafe_0000), 1);
}

#[test]
fn empty_pool_state() {
    let p = TypePool::new();
    assert!(p.is_empty());
    assert_eq!(p.len(), 0);
}
