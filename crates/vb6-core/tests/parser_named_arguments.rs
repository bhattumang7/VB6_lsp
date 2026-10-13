//! A named argument's name may be any word, including words that are keywords
//! elsewhere; the name labels the callee's parameter, never a variable use.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::session::Session;

fn body(stmts_src: &str) -> (usize, Vec<String>) {
    let src = format!("Sub T()\n{stmts_src}\nEnd Sub\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Standard);
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    let ExprNode::ProcDecl { body, .. } = arena.get(nodes[0]) else { panic!("ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("Block") };
    (diags, stmts.iter().map(|s| render(&ctx, &arena, *s)).collect())
}

fn render(ctx: &ScannerContext, arena: &ExprArena, id: NodeId) -> String {
    let name = |s: u32| ctx.symbol(s as usize).name.clone();
    let list = |a: NodeId| match arena.get(a) {
        ExprNode::ArgList { args } => args.iter().map(|x| render(ctx, arena, *x)).collect::<Vec<_>>().join(" "),
        _ => panic!("ArgList"),
    };
    match arena.get(id) {
        ExprNode::NameRef { sym, .. } => name(*sym),
        ExprNode::Literal { lit } => format!("{lit:?}"),
        ExprNode::NamedArg { name: n, value } => format!("({}:={})", name(*n), render(ctx, arena, *value)),
        ExprNode::MemberAccess { base, member, .. } => format!("(. {} {})", render(ctx, arena, *base), name(*member)),
        ExprNode::Call { func, args } => format!("(call {} [{}])", render(ctx, arena, *func), list(*args)),
        ExprNode::CallStmt { callee, args } => format!("(stmt {} [{}])", render(ctx, arena, *callee), list(*args)),
        ExprNode::Assign { target, value } => format!("(= {} {})", render(ctx, arena, *target), render(ctx, arena, *value)),
        _ => panic!("unexpected node"),
    }
}

#[test]
fn keyword_named_arguments_in_a_call_statement() {
    assert_eq!(
        body("c.ListItems.Add key:=s.key, text:=s.text"),
        (0, vec!["(stmt (. (. c ListItems) Add) [(key:=(. s key)) (Text:=(. s Text))])".to_string()])
    );
}

#[test]
fn keyword_named_arguments_with_parentheses() {
    assert_eq!(
        body("x = f(Name:=1, Key:=2, Type:=3)"),
        (0, vec![
            "(= x (call F [(Name:=Int(1)) (Key:=Int(2)) (Type:=Int(3))]))".to_string()
        ])
    );
}

#[test]
fn plain_identifier_named_arguments_are_unchanged() {
    assert_eq!(
        body("Foo first:=1, second:=2"),
        (0, vec!["(stmt Foo [(first:=Int(1)) (second:=Int(2))])".to_string()])
    );
}

#[test]
fn argument_name_is_not_a_variable_use() {
    let src = "Option Explicit\nSub T()\n    Dim c As Object\n    c.Add key:=1, text:=2\nEnd Sub\n";
    let s = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
    assert_eq!(s.diagnostics(0), vec![]);
}
