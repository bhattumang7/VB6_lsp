//! The procedures of one property agree on their index parameters and, for a
//! getter and a `Property Let`, on the value type. The disagreement is reported
//! at the name of the later procedure.

mod common;
use common::{offset_of, parse};

const INCONSISTENT: u32 = 0x9c83;

fn pair(first: &str, second: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("{first}\nEnd Property\n{second}\nEnd Property\n");
    let found = parse(&src).found;
    (src, found)
}

fn second_name(src: &str) -> u32 {
    offset_of(src, "P(", 1)
}

#[test]
fn a_getter_and_a_let_agree_on_the_value_type() {
    for (get, let_) in [
        ("Property Get P() As Long", "Property Let P(v As Long)"),
        ("Property Get P() As Long", "Property Let P(ByVal v As Long)"),
        ("Property Get P() As String", "Property Let P(v As String)"),
        ("Property Get P() As Variant", "Property Let P(v)"),
        ("Property Get P()", "Property Let P(v As Variant)"),
        ("Property Get P()", "Property Let P(v)"),
        ("Property Get P() As Object", "Property Let P(v As Variant)"),
        ("Property Get P() As Object", "Property Let P(v)"),
        ("Property Get P() As Foo", "Property Let P(v As Variant)"),
        ("Property Get P() As Foo", "Property Let P(v As Foo)"),
        ("Property Get P() As Object", "Property Let P(v As Object)"),
        ("Property Get P() As Long()", "Property Let P(v() As Long)"),
    ] {
        assert_eq!(pair(get, let_).1, vec![], "{get} / {let_}");
        assert_eq!(pair(let_, get).1, vec![], "{let_} / {get}");
    }
}

#[test]
fn a_getter_and_a_let_with_different_value_types_disagree_at_the_later_one() {
    for (get, let_) in [
        ("Property Get P() As Long", "Property Let P(v As String)"),
        ("Property Get P() As Long", "Property Let P(v As Variant)"),
        ("Property Get P() As Long", "Property Let P(v)"),
        ("Property Get P() As Variant", "Property Let P(v As Long)"),
        ("Property Get P() As Variant", "Property Let P(v As Object)"),
        ("Property Get P() As Object", "Property Let P(v As Foo)"),
        ("Property Get P() As Foo", "Property Let P(v As Object)"),
        ("Property Get P() As Long()", "Property Let P(v As Long)"),
        ("Property Get P() As Long()", "Property Let P(v() As Variant)"),
        ("Property Get P() As Long()", "Property Let P(v() As Single)"),
        ("Property Get P() As Variant", "Property Let P(v() As Long)"),
        ("Property Get P() As Long", "Property Let P(v() As Long)"),
    ] {
        let (src, found) = pair(get, let_);
        assert_eq!(found, vec![(INCONSISTENT, second_name(&src), 1)], "{get} / {let_}");
        let (src, found) = pair(let_, get);
        assert_eq!(found, vec![(INCONSISTENT, second_name(&src), 1)], "{let_} / {get}");
    }
}

#[test]
fn a_set_agrees_with_any_getter_and_let_value_but_not_with_other_indexes() {
    for get in ["Property Get P() As Long", "Property Get P() As String", "Property Get P() As Object"] {
        assert_eq!(pair(get, "Property Set P(v As Object)").1, vec![], "{get}");
    }
    assert_eq!(pair("Property Let P(v As Long)", "Property Set P(v As Object)").1, vec![]);
}

#[test]
fn index_parameters_must_match_in_count_name_type_and_pass_mode() {
    let agree = [
        ("Property Get P(i As Long) As Long", "Property Let P(i As Long, v As Long)"),
        ("Property Get P(i As Long) As Long", "Property Let P(Optional i As Long, v As Long)"),
        ("Property Get P(i) As Long", "Property Let P(i As Variant, v As Long)"),
        ("Property Get P(ByVal i As Long) As Long", "Property Let P(ByVal i As Long, v As Long)"),
        ("Property Get P(i() As Long) As Long", "Property Let P(i() As Long, v As Long)"),
        ("Property Get P(i As Long, j As Long) As Long", "Property Set P(i As Long, j As Long, v As Object)"),
    ];
    for (a, b) in agree {
        assert_eq!(pair(a, b).1, vec![], "{a} / {b}");
    }
    let disagree = [
        ("Property Get P(i As Long) As Long", "Property Let P(v As Long)"),
        ("Property Get P(i As Long) As Long", "Property Let P(j As Long, v As Long)"),
        ("Property Get P(i As Integer) As Long", "Property Let P(i As Long, v As Long)"),
        ("Property Get P(ByVal i As Long) As Long", "Property Let P(i As Long, v As Long)"),
        ("Property Get P(i As Long) As Long", "Property Set P(ByVal i As Long, v As Object)"),
        ("Property Get P(i As Long, j As Long) As Long", "Property Let P(i As Long, v As Long)"),
        ("Property Let P(i As Long, v As Long)", "Property Set P(v As Object)"),
    ];
    for (a, b) in disagree {
        let (src, found) = pair(a, b);
        assert_eq!(found.len(), 1, "{a} / {b}: {found:?}");
        assert_eq!(found[0].0, INCONSISTENT, "{a} / {b}");
        assert_eq!(found[0].1, second_name(&src), "{a} / {b}");
    }
}

#[test]
fn a_third_procedure_is_checked_against_each_earlier_one() {
    let src = "Property Get P() As Long\nEnd Property\nProperty Let P(v As Long)\nEnd Property\nProperty Set P(i As Long, v As Object)\nEnd Property\n";
    let found = parse(src).found;
    assert_eq!(found, vec![(INCONSISTENT, offset_of(src, "P(i", 0), 1)]);
}

#[test]
fn procedures_of_different_names_are_independent() {
    let src = "Property Get P() As Long\nEnd Property\nProperty Let Q(v As String)\nEnd Property\n";
    assert_eq!(parse(src).found, vec![]);
}
