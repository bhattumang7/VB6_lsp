use super::*;
use std::collections::HashSet;

#[test]
fn no_duplicate_or_empty_entries() {
    let mut seen = HashSet::new();
    for &name in BUILTINS {
        assert!(!name.is_empty(), "empty builtin name");
        assert_eq!(name, name.to_ascii_lowercase(), "builtin `{name}` must be lowercase");
        assert!(seen.insert(name), "duplicate builtin entry: {name}");
    }
}

#[test]
fn recognises_runtime_intrinsics_previously_missing() {
    // Standard VB6 intrinsics that were absent before and would have
    // produced a spurious "Sub or Function not defined".
    for name in [
        "dir", "dir$", "randomize", "rgb", "qbcolor", "imestatus", "cvdate",
        "pmt", "ppmt", "ipmt", "fv", "pv", "nper", "npv", "irr", "mirr",
        "ddb", "sln", "syd", "rate",
        "formatnumber", "formatcurrency", "formatpercent", "formatdatetime",
    ] {
        assert!(is_builtin(name), "expected `{name}` to be a recognised builtin");
    }
}

#[test]
fn lookup_is_order_independent() {
    // `spc` sits after `split` in source order; a binary search would miss
    // it. The set-based lookup must still find it.
    assert!(is_builtin("spc"));
    assert!(is_builtin("split"));
    assert!(!is_builtin("definitelynotabuiltin"));
}
