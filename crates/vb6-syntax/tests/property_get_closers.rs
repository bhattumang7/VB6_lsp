//! A property getter is closed by `End Sub`, `End Function` or `End Property`,
//! and may be left with `Exit Sub`, `Exit Function` or `Exit Property`; the
//! other property procedures keep the strict rules.

mod common;
use common::{offset_of, parse};

const EXPECTED_END_PROPERTY: u32 = 0x9cbe;
const EXIT_SUB_NOT_ALLOWED: u32 = 0x9ca3;
const EXIT_FUNCTION_NOT_ALLOWED: u32 = 0x9ca2;

#[test]
fn a_getter_accepts_each_closer() {
    for closer in ["End Property", "End Sub", "End Function"] {
        let src = format!("Property Get P() As Long\n{closer}\n");
        assert_eq!(parse(&src).found, vec![], "{closer}");
    }
    let src = "Property Get P() As Long\nEnd If\n";
    assert_eq!(parse(src).found[0].0, 40087);
}

#[test]
fn a_getter_accepts_each_exit() {
    for exit in ["Exit Property", "Exit Sub", "Exit Function"] {
        let src = format!("Property Get P() As Long\n    {exit}\nEnd Property\n");
        assert_eq!(parse(&src).found, vec![], "{exit}");
    }
}

#[test]
fn a_getter_still_rejects_loop_exits_outside_their_loops() {
    let src = "Property Get P() As Long\n    Exit Do\nEnd Property\n";
    assert_eq!(parse(src).found[0].1, offset_of(src, "Exit", 0));
}

#[test]
fn let_and_set_keep_the_strict_closer_and_exit_rules() {
    for head in ["Property Let P(v)", "Property Set P(v)"] {
        for closer in ["End Sub", "End Function"] {
            let src = format!("{head}\n{closer}\n");
            assert_eq!(parse(&src).found, vec![(EXPECTED_END_PROPERTY, offset_of(&src, "End", 0), 3)], "{head} {closer}");
        }
        let src = format!("{head}\n    Exit Sub\nEnd Property\n");
        assert_eq!(parse(&src).found, vec![(EXIT_SUB_NOT_ALLOWED, offset_of(&src, "Exit", 0), 4)], "{head}");
        let src = format!("{head}\n    Exit Function\nEnd Property\n");
        assert_eq!(parse(&src).found, vec![(EXIT_FUNCTION_NOT_ALLOWED, offset_of(&src, "Exit", 0), 4)], "{head}");
    }
}
