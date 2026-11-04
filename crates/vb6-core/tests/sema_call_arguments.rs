//! Argument checks on calls to procedures of the same module.

use vb6_core::session::Session;

const NAMED_NOT_FOUND: u32 = 448;
const NOT_OPTIONAL: u32 = 449;
const WRONG_NUMBER: u32 = 450;

fn codes(src: &str) -> Vec<(u32, u32, u32)> {
    let s = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
    s.diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

fn at(src: &str, ctx: &str, member: &str) -> u32 {
    (src.find(ctx).unwrap() + ctx.find(member).unwrap()) as u32
}

const DEFS: &str = "Sub P(a As Long, Optional b As Long)\nEnd Sub\nFunction F(x As Long) As Long\nEnd Function\nSub Q(ParamArray rest())\nEnd Sub\nSub R()\nEnd Sub\n";

#[test]
fn too_many_arguments() {
    let src = format!("{DEFS}Sub Run()\n    P 1, 2, 3\n    R 1\n    x = F(1, 2)\nEnd Sub\n");
    assert_eq!(
        codes(&src),
        vec![
            (WRONG_NUMBER, at(&src, "P 1, 2, 3", "P"), 1),
            (WRONG_NUMBER, at(&src, "R 1", "R"), 1),
            (WRONG_NUMBER, at(&src, "F(1, 2)", "F"), 1),
        ]
    );
}

#[test]
fn a_function_returning_an_object_is_indexed_by_the_extra_arguments() {
    let src = "Function Items() As Object\nEnd Function\nFunction Any1()\nEnd Function\nFunction Lst() As Long()\nEnd Function\nSub Run()\n    x = Items(\"k\")\n    Items(\"k\").Add 1\n    y = Any1(2)\n    z = Lst(1)\nEnd Sub\n";
    assert_eq!(codes(src), vec![]);
}

#[test]
fn a_function_with_a_scalar_result_has_a_fixed_argument_count() {
    let src = "Function N() As Long\nEnd Function\nSub Run()\n    x = N(1)\nEnd Sub\n";
    assert_eq!(codes(src), vec![(WRONG_NUMBER, at(src, "N(1)", "N"), 1)]);
}

#[test]
fn a_required_parameter_needs_an_argument() {
    let src = format!("{DEFS}Sub Run()\n    P\n    P , 2\n    x = F()\n    Call P\n    x = F\nEnd Sub\n");
    let found: Vec<u32> = codes(&src).iter().map(|d| d.0).collect();
    assert_eq!(found, vec![NOT_OPTIONAL; 5]);
}

#[test]
fn a_named_argument_must_name_a_parameter() {
    let src = format!("{DEFS}Sub Run()\n    P 1, zz:=2\nEnd Sub\n");
    let found = codes(&src);
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].0, NAMED_NOT_FOUND);
}

#[test]
fn valid_calls_stay_silent() {
    let src = format!(
        "{DEFS}Sub Run()\n    P 1\n    P 1, 2\n    P b:=2, a:=1\n    P A:=1\n    Call P(1)\n    Q\n    Q 1, 2, 3\n    R\n    Call R\n    x = F(1)\n    x = F(x:=3)\nEnd Sub\n"
    );
    assert_eq!(codes(&src), vec![]);
}

#[test]
fn messages_are_the_compiler_texts() {
    let src = format!("{DEFS}Sub Run()\n    R 1\n    P\n    P 1, zz:=2\nEnd Sub\n");
    let s = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
    let mut msgs: Vec<_> = s.diagnostics(0).iter().map(|d| d.message().unwrap_or_default()).collect();
    msgs.sort();
    assert_eq!(
        msgs,
        vec![
            "Argument not optional".to_string(),
            "Named argument not found".to_string(),
            "Wrong number of arguments or invalid property assignment".to_string(),
        ]
    );
}
