//! Built-in functions with a fixed argument grammar: the accepted shapes build
//! a call of the right arity; the others report one diagnostic at the name.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, NodeId};

fn rhs(p: &Parsed) -> NodeId {
    let body = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(body[0]) else { panic!("expected an assignment") };
    *value
}

fn call_parts(p: &Parsed, id: NodeId) -> (String, Vec<NodeId>) {
    let ExprNode::Call { func, args } = p.node(id) else { panic!("expected a call") };
    let ExprNode::NameRef { sym, .. } = p.node(*func) else { panic!("expected a name") };
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("expected arguments") };
    (p.name(*sym).to_string(), args.clone())
}

fn src(expr: &str) -> String {
    format!("Sub S()\n    x = {expr}\nEnd Sub\n")
}

#[test]
fn accepted_arities_build_calls_with_those_arguments() {
    for (text, name, count) in [
        ("Len(a)", "Len", 1), ("Abs(1)", "Abs", 1), ("InStr(a, b)", "InStr", 2), ("InStr(1, a, b, 1)", "InStr", 4),
        ("StrComp(a, b)", "StrComp", 2), ("StrComp(a, b, 1)", "StrComp", 3), ("String(3, a)", "String", 2),
        ("UBound(a)", "UBound", 1), ("UBound(a, 2)", "UBound", 2), ("LBound(o.p)", "LBound", 1),
    ] {
        let p = parse(&src(text));
        assert_eq!(p.found, vec![], "{text}");
        let (n, args) = call_parts(&p, rhs(&p));
        assert_eq!((n.as_str(), args.len()), (name, count), "{text}");
    }
}

#[test]
fn wrong_arity_is_one_diagnostic_at_the_name() {
    for (text, name) in [
        ("Len()", "Len"), ("Len(a, b)", "Len"), ("InStr(a)", "InStr"), ("InStr(1, a, b, 1, 2)", "InStr"),
        ("StrComp(a)", "StrComp"), ("String(3)", "String"), ("UBound(a, 1, 2)", "UBound"),
    ] {
        let s = src(text);
        let p = parse(&s);
        assert_eq!(p.found, vec![(0x9c70, offset_of(&s, name, 0), name.len() as u32)], "{text}");
    }
}

#[test]
fn empty_and_named_arguments_are_rejected() {
    for text in ["Len(, a)", "Abs(1, )", "Len(a:=1)"] {
        let s = src(text);
        let p = parse(&s);
        assert!(!p.found.is_empty(), "{text}");
        let name_at = offset_of(&s, text.split('(').next().unwrap(), 0);
        assert!(p.found.iter().any(|d| d.0 == 0x9c70 && d.1 == name_at), "{text}: {:?}", p.found);
    }
}

#[test]
fn the_array_of_bound_functions_must_be_a_variable() {
    let s = src("UBound(1)");
    let p = parse(&s);
    assert_eq!(p.found, vec![(0x9c70, offset_of(&s, "UBound", 0), 6)]);
}

#[test]
fn bare_names_need_brackets_except_date() {
    let s = src("Len");
    assert_eq!(parse(&s).found, vec![(0x9c70, offset_of(&s, "Len", 0), 3)]);
    let p = parse(&src("Date"));
    assert_eq!(p.found, vec![]);
    assert!(matches!(p.node(rhs(&p)), ExprNode::NameRef { .. }));
}

#[test]
fn array_is_a_function_that_takes_any_number_of_arguments() {
    let p = parse(&src("Array(1, 2, 3)"));
    assert_eq!(p.found, vec![]);
    let (n, args) = call_parts(&p, rhs(&p));
    assert_eq!((n.as_str(), args.len()), ("Array", 3));
    assert!(matches!(p.node(args[0]), ExprNode::Literal { lit: AstLit::Int(1) }));
    let s = src("Array");
    assert_eq!(parse(&s).found, vec![(0x9c70, offset_of(&s, "Array", 0), 5)]);
}

#[test]
fn address_of_is_only_a_call_argument() {
    let p = parse("Sub S()\n    x = F(AddressOf G)\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let (_, args) = call_parts(&p, rhs(&p));
    let ExprNode::AddressOf { operand } = p.node(args[0]) else { panic!("expected AddressOf") };
    let ExprNode::NameRef { sym, .. } = p.node(*operand) else { panic!("expected a name") };
    assert_eq!(p.name(*sym), "G");
    for text in ["x = AddressOf G", "x = 1 + AddressOf G"] {
        let s = format!("Sub S()\n    {text}\nEnd Sub\n");
        let p = parse(&s);
        assert_eq!(p.found.len(), 1, "{text}");
        assert_eq!(p.found[0].1, offset_of(&s, "AddressOf", 0), "{text}");
    }
}

#[test]
fn nothing_joins_to_a_group() {
    let s = src("(a).b");
    let p = parse(&s);
    assert_eq!(p.found.len(), 1);
    assert_eq!(p.found[0].1, offset_of(&s, ".b", 0));
}
