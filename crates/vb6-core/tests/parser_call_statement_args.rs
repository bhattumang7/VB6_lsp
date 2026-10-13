//! Call statements without `Call`: a parenthesised group after the callee
//! begins the first argument expression when the statement goes on; omitted
//! arguments are `MissingArg`; graphics coordinate pairs keep their form.

use vb6_core::frontend::ast::{BinOpKind, ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

/// (diagnostics, rendered statements of `Sub T`).
fn body(stmts_src: &str) -> (usize, Vec<String>) {
    let src = format!("Sub T()\n{stmts_src}\nEnd Sub\n");
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Form);
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    drop(parser);
    let ExprNode::ProcDecl { body, .. } = arena.get(nodes[0]) else { panic!("expected ProcDecl") };
    let ExprNode::Block { stmts } = arena.get(*body) else { panic!("expected Block") };
    (diags, stmts.iter().map(|s| render(&ctx, &arena, *s)).collect())
}

fn render(ctx: &ScannerContext, arena: &ExprArena, id: NodeId) -> String {
    let name = |s: u32| ctx.symbol(s as usize).name.clone();
    let list = |a: NodeId| match arena.get(a) {
        ExprNode::ArgList { args } => args.iter().map(|x| render(ctx, arena, *x)).collect::<Vec<_>>().join(" "),
        _ => panic!("expected ArgList"),
    };
    match arena.get(id) {
        ExprNode::Me => "Me".into(),
        ExprNode::NameRef { sym, .. } => name(*sym),
        ExprNode::Literal { lit } => format!("{lit:?}"),
        ExprNode::MissingArg => "_".into(),
        ExprNode::Paren { inner } => format!("(paren {})", render(ctx, arena, *inner)),
        ExprNode::MemberAccess { base, member, .. } => format!("(. {} {})", render(ctx, arena, *base), name(*member)),
        ExprNode::BinOp { op, lhs, rhs } => {
            let o = match op {
                BinOpKind::Sub => "-",
                BinOpKind::Div => "/",
                BinOpKind::Cat => "&",
                BinOpKind::Add => "+",
                other => panic!("unexpected operator {other:?}"),
            };
            format!("({o} {} {})", render(ctx, arena, *lhs), render(ctx, arena, *rhs))
        }
        ExprNode::Call { func, args } => format!("(call {} [{}])", render(ctx, arena, *func), list(*args)),
        ExprNode::CallStmt { callee, args } => format!("(stmt {} [{}])", render(ctx, arena, *callee), list(*args)),
        ExprNode::Assign { target, value } => format!("(= {} {})", render(ctx, arena, *target), render(ctx, arena, *value)),
        _ => panic!("unexpected node"),
    }
}

fn one(src: &str) -> String {
    let (diags, stmts) = body(src);
    assert_eq!(diags, 0, "diagnostics for {src:?}");
    assert_eq!(stmts.len(), 1, "one statement for {src:?}");
    stmts.into_iter().next().unwrap()
}

#[test]
fn parenthesised_group_starts_an_operator_expression() {
    assert_eq!(
        one("oPic.Move (nWidth / 2) - (nOther / 2), nTop"),
        "(stmt (. oPic Move) [(- (paren (/ nWidth Int(2))) (paren (/ nOther Int(2)))) nTop])"
    );
}

#[test]
fn parenthesised_group_after_a_line_continuation() {
    assert_eq!(
        one("oPic.Move _\n    (nWidth / 2) - (nOther / 2), _\n    nTop"),
        "(stmt (. oPic Move) [(- (paren (/ nWidth Int(2))) (paren (/ nOther Int(2)))) nTop])"
    );
}

#[test]
fn parenthesised_first_argument_followed_by_more_arguments() {
    assert_eq!(one("MsgBox (\"Saved\"), vbInformation"), "(stmt MsgBox [(paren Str(\"Saved\")) vbInformation])");
}

#[test]
fn parenthesised_group_with_concatenation() {
    assert_eq!(one("Log (sName) & \"!\""), "(stmt Log [(& (paren sName) Str(\"!\"))])");
}

#[test]
fn single_parenthesised_argument_alone_stays_a_call() {
    assert_eq!(one("MsgBox (\"Saved\")"), "(call MsgBox [Str(\"Saved\")])");
}

#[test]
fn indexed_assignment_target_is_unchanged() {
    assert_eq!(one("aValues(1) = nTotal + 1"), "(= (call aValues [Int(1)]) (+ nTotal Int(1)))");
}

#[test]
fn omitted_arguments_in_a_call_statement() {
    assert_eq!(one("MsgBox \"Text\", , \"Title\""), "(stmt MsgBox [Str(\"Text\") _ Str(\"Title\")])");
    assert_eq!(one("Foo , nSecond"), "(stmt Foo [_ nSecond])");
    assert_eq!(one("MsgBox (\"Text\"), , \"Title\""), "(stmt MsgBox [(paren Str(\"Text\")) _ Str(\"Title\")])");
}

#[test]
fn graphics_line_coordinate_pairs_keep_their_form() {
    assert_eq!(
        one("Me.Line (0, 0)-(10, 10), vbRed"),
        "(stmt (call (. Me Line) [Int(0) Int(0)]) [Int(10) Int(10) vbRed])"
    );
}

#[test]
fn graphics_pset_coordinate_pair_keeps_its_form() {
    assert_eq!(one("PSet (nX, nY), vbRed"), "(stmt (call PSet [nX nY]) [vbRed])");
}
