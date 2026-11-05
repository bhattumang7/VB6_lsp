//! Argument values: `ByVal` and `AddressOf` may start the value of a named argument,
//! `ByVal` cannot be followed by `AddressOf` and has no place in a built-in function with
//! a fixed argument list, `AddressOf` takes a bracketed name, and `TypeOf` does not take
//! `Array(...)`.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

fn in_sub(body: &str) -> (String, common::Parsed) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let p = parse(&src);
    (src, p)
}

#[test]
fn a_named_argument_value_may_start_with_byval_or_addressof() {
    for body in [
        "T a:=AddressOf Bar",
        "T a:=1, b:=AddressOf Bar",
        "x = Dcl(a:=AddressOf Bar)",
        "T a:=ByVal 1",
        "x = Bar(a:=ByVal 1)",
        "T ByVal 1, a:=ByVal 2",
    ] {
        let (_, p) = in_sub(body);
        assert_eq!(p.found, vec![], "{body}");
    }
}

#[test]
fn a_named_argument_keeps_its_addressof_value() {
    let (_, p) = in_sub("x = Dcl(a:=AddressOf Bar)");
    let b = p.proc_body(0);
    let ExprNode::Assign { value, .. } = p.node(b[0]) else { panic!("not an assignment") };
    let ExprNode::Call { args, .. } = p.node(*value) else { panic!("not a call") };
    let ExprNode::ArgList { args } = p.node(*args) else { panic!("no args") };
    let ExprNode::NamedArg { value, .. } = p.node(args[0]) else { panic!("not named") };
    assert!(matches!(p.node(*value), ExprNode::AddressOf { .. }));
}

#[test]
fn byval_cannot_be_followed_by_addressof() {
    for (body, at) in [
        ("x = Dcl(ByVal AddressOf Bar)", "AddressOf"),
        ("x = Dcl(a:=ByVal AddressOf Bar)", "AddressOf"),
        ("T ByVal AddressOf Bar", "AddressOf"),
    ] {
        let (src, p) = in_sub(body);
        assert_eq!(p.found[0].1, offset_of(&src, at, 0), "{body}");
    }
}

#[test]
fn byval_is_not_taken_by_built_in_functions_with_fixed_arguments() {
    for call in ["Len(ByVal s)", "CInt(ByVal s)", "StrComp(ByVal s, s)", "InStr(ByVal s, s)", "UBound(ByVal s)"] {
        let (src, p) = in_sub(&format!("x = {call}"));
        assert!(!p.found.is_empty(), "{call}");
        assert_eq!(p.found[0].1, offset_of(&src, "ByVal", 0), "{call}");
    }
}

#[test]
fn addressof_takes_a_bracketed_name() {
    for body in ["x = Dcl(AddressOf [Bar])", "x = Dcl(AddressOf [Bar].Baz)", "x = Dcl(AddressOf Bar.Baz)"] {
        let (_, p) = in_sub(body);
        assert_eq!(p.found, vec![], "{body}");
    }
}

#[test]
fn typeof_does_not_take_array() {
    let (src, p) = in_sub("If TypeOf Array(1) Is Object Then x = 1");
    assert_eq!(p.found[0].1, offset_of(&src, "Array", 0));
    let (_, p) = in_sub("If TypeOf o Is Object Then x = 1");
    assert_eq!(p.found, vec![]);
}

#[test]
fn call_site_byval_is_recorded_for_the_argument() {
    use vb6_syntax::frontend::ast::ExprArena;
    use vb6_syntax::frontend::parser::Parser;
    use vb6_syntax::frontend::scanner::ScannerContext;
    let src = "Sub S()\n T ByVal x, y\n z = F(ByVal p, q, a:=ByVal r)\nEnd Sub\n";
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    parser.parse_module(&mut arena);
    let names: Vec<u32> = parser
        .by_val_args
        .iter()
        .map(|n| match arena.get(*n) {
            ExprNode::NameRef { sym, .. } => *sym,
            _ => panic!("not a name"),
        })
        .collect();
    drop(parser);
    let spelled: Vec<String> = names.iter().map(|s| ctx.sym_name(*s as usize).to_string()).collect();
    assert_eq!(spelled, vec!["x", "p", "r"]);
}
