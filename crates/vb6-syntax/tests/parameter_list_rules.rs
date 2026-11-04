//! Shape rules for parameter lists, property value parameters and events.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprNode, ProcKind};
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::scanner::ScannerContext;

const SYNTAX_ERROR: u32 = 0x9c70;
const EXPECTED: u32 = 0x9c6f;
const ARRAY_MUST_BE_BYREF: u32 = 0x9ca0;
const PROPERTY_ARGUMENT_REQUIRED: u32 = 0x9cbd;
const OPTIONAL_ARRAY: u32 = 0x9ccf;
const PARAM_ARRAY_NOT_ARRAY: u32 = 0x9cd0;
const INVALID_EVENT_NAME: u32 = 0xdefb;
const PARAM_ARRAY_WITH_OPTIONAL: u32 = 0xdf00;

/// `(code, label, start, len)` of every diagnostic of an object module.
fn labelled(src: &str) -> Vec<(u32, Option<&'static str>, u32, u32)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    parser.parse_module(&mut arena);
    parser
        .diagnostics
        .items()
        .iter()
        .map(|d| (d.code, d.label, d.span.start, d.span.len))
        .collect()
}

fn sub(params: &str) -> String {
    format!("Sub S({params})\nEnd Sub\n")
}

#[test]
fn well_formed_lists_are_clean() {
    for params in [
        "",
        "a, b As Long, ByVal c As String, ByRef d() As Long",
        "a, Optional b, Optional ByVal c As Long = 5",
        "a, ParamArray rest() As Variant",
        "ByVal a%, b$",
    ] {
        let src = sub(params);
        assert_eq!(parse(&src).found, vec![], "for ({params})");
    }
}

#[test]
fn a_required_parameter_cannot_follow_an_optional_one() {
    let src = sub("Optional xx, yy");
    let p = parse(&src);
    assert_eq!(p.found, vec![(EXPECTED, offset_of(&src, "yy", 0), 2)]);
    assert_eq!(labelled(&src), vec![(EXPECTED, Some("Optional"), offset_of(&src, "yy", 0), 2)]);
    // The procedure and both parameters are still built.
    let ExprNode::ProcDecl { params: Some(list), .. } = p.node(p.top[0]) else { panic!("expected Sub") };
    let ExprNode::ArgList { args } = p.node(*list) else { panic!("expected ArgList") };
    assert_eq!(args.len(), 2);
}

#[test]
fn param_array_must_be_last_and_an_array() {
    let src = sub("ParamArray xx(), yy");
    assert_eq!(labelled(&src), vec![(EXPECTED, Some(")"), offset_of(&src, "xx", 0), 2)]);

    let src = sub("ParamArray xx");
    assert_eq!(parse(&src).found, vec![(PARAM_ARRAY_NOT_ARRAY, offset_of(&src, "xx", 0), 2)]);

    let src = sub("ParamArray xx(), ParamArray yy()");
    assert_eq!(labelled(&src), vec![(EXPECTED, Some(")"), offset_of(&src, "xx", 0), 2)]);
}

#[test]
fn optional_and_param_array_do_not_mix_in_the_wrong_order() {
    let src = sub("Optional xx, ParamArray yy()");
    assert_eq!(parse(&src).found, vec![(PARAM_ARRAY_WITH_OPTIONAL, offset_of(&src, "yy", 0), 2)]);
    let src = sub("ParamArray xx(), Optional yy");
    assert_eq!(labelled(&src), vec![(EXPECTED, Some(")"), offset_of(&src, "xx", 0), 2)]);
}

#[test]
fn array_parameters_cannot_be_byval_or_optional() {
    let src = sub("ByVal xx() As Long");
    assert_eq!(parse(&src).found, vec![(ARRAY_MUST_BE_BYREF, offset_of(&src, "xx", 0), 2)]);
    let src = sub("Optional xx() As Long");
    assert_eq!(parse(&src).found, vec![(OPTIONAL_ARRAY, offset_of(&src, "xx", 0), 2)]);
    assert_eq!(parse(&sub("ByRef xx() As Long")).found, vec![]);
}

#[test]
fn only_the_first_violation_of_a_list_is_reported() {
    let src = sub("Optional xx, yy, zz, ParamArray ww");
    assert_eq!(parse(&src).found.len(), 1);
}

#[test]
fn property_let_and_set_need_a_value_parameter() {
    for (head, kind) in [("Property Let Prp", ProcKind::PropLet), ("Property Set Prp", ProcKind::PropSet)] {
        for params in ["", "()"] {
            let src = format!("{head}{params}\nEnd Property\n");
            let p = parse(&src);
            assert_eq!(
                p.found,
                vec![(PROPERTY_ARGUMENT_REQUIRED, offset_of(&src, "Prp", 0), 3)],
                "for {src:?}"
            );
            let ExprNode::ProcDecl { kind: got, .. } = p.node(p.top[0]) else { panic!("expected property") };
            assert_eq!(*got, kind);
        }
        let ty = if kind == ProcKind::PropSet { "Object" } else { "Long" };
        let src = format!("{head}(v As {ty})\nEnd Property\n");
        assert_eq!(parse(&src).found, vec![], "for {src:?}");
    }
    // Property Get takes none.
    assert_eq!(parse("Property Get Prp() As Long\nEnd Property\n").found, vec![]);
}

#[test]
fn property_let_value_parameter_may_follow_optional_ones() {
    let src = "Property Let Prp(Optional ii, vv As Long)\nEnd Property\n";
    assert_eq!(parse(src).found, vec![]);
}

#[test]
fn events_take_no_optional_or_param_array_and_no_underscore_name() {
    let src = "Event E(Optional xx)\n";
    assert_eq!(labelled(src), vec![(SYNTAX_ERROR, None, offset_of(src, "xx", 0), 2)]);
    let src = "Event E(ParamArray xx())\n";
    assert_eq!(labelled(src), vec![(SYNTAX_ERROR, None, offset_of(src, "xx", 0), 2)]);
    let src = "Event Bad_Name(xx As Long)\n";
    assert_eq!(labelled(src), vec![(INVALID_EVENT_NAME, None, offset_of(src, "Bad_Name", 0), 8)]);
    assert_eq!(labelled("Event Good(xx As Long, yy)\n"), vec![]);
}

#[test]
fn param_array_element_type_must_be_variant() {
    for ok in ["Sub S(ParamArray p() As Variant)\nEnd Sub\n", "Sub S(ParamArray p())\nEnd Sub\n"] {
        assert_eq!(labelled(ok), vec![], "{ok:?}");
    }
    for ty in ["Long", "String", "Object", "MyType"] {
        let src = format!("Sub S(ParamArray p() As {ty})\nEnd Sub\n");
        let found = labelled(&src);
        assert_eq!(found.len(), 1, "{ty}");
        assert_eq!(found[0].0, 0x9cd0, "{ty}");
        assert_eq!(found[0].2, src.find("p()").unwrap() as u32, "{ty}");
    }
}
