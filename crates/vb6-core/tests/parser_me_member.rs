//! `Me` takes member access and calls like any object, in statements,
//! `Set` targets and expressions; `Me.member` is a reference to a public member
//! of the class itself.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::session::Session;

/// (diagnostics, the body statements of `Sub T` rendered).
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
        ExprNode::MemberAccess { base, member, bang } => {
            format!("({} {} {})", if *bang { "!" } else { "." }, render(ctx, arena, *base), name(*member))
        }
        ExprNode::Call { func, args } => format!("(call {} [{}])", render(ctx, arena, *func), list(*args)),
        ExprNode::CallStmt { callee, args } => format!("(stmt {} [{}])", render(ctx, arena, *callee), list(*args)),
        ExprNode::Assign { target, value } => format!("(= {} {})", render(ctx, arena, *target), render(ctx, arena, *value)),
        ExprNode::SetAssign { target, value } => {
            format!("(set {} {})", render(ctx, arena, *target), render(ctx, arena, *value))
        }
        _ => panic!("unexpected node"),
    }
}

#[test]
fn me_method_call_statement_with_arguments() {
    assert_eq!(
        body("Me.SetStatusText \"\"\nMe.SetStatusLabel \"a\", \"b\""),
        (0, vec![
            "(stmt (. Me SetStatusText) [Str(\"\")])".to_string(),
            "(stmt (. Me SetStatusLabel) [Str(\"a\") Str(\"b\")])".to_string(),
        ])
    );
}

#[test]
fn me_method_call_statement_without_arguments() {
    assert_eq!(body("Me.Refresh"), (0, vec!["(. Me Refresh)".to_string()]));
}

#[test]
fn set_me_property() {
    assert_eq!(
        body("Set Me.Icon = frmOther.Icon"),
        (0, vec!["(set (. Me Icon) (. frmOther Icon))".to_string()])
    );
}

#[test]
fn me_member_chain_in_an_expression() {
    assert_eq!(
        body("bShown = Me.Controls(\"cmdOK\").Visible"),
        (0, vec!["(= bShown (. (call (. Me Controls) [Str(\"cmdOK\")]) Visible))".to_string()])
    );
}

#[test]
fn me_bang_member() {
    assert_eq!(body("sValue = Me!Field"), (0, vec!["(= sValue (! Me Field))".to_string()]));
}

#[test]
fn bare_me_argument() {
    assert_eq!(body("Show Me"), (0, vec!["(stmt Show [Me])".to_string()]));
}

fn text(src: &str, start: u32, len: u32) -> &str {
    &src[start as usize..(start + len) as usize]
}

#[test]
fn me_member_references_the_public_procedure_of_the_class() {
    let src = "Public Sub Refresh()\nEnd Sub\nPrivate Sub Helper()\nEnd Sub\nSub T()\n    Me.Refresh\n    Me.Helper\nEnd Sub\n";
    let s = Session::from_sources(vec![("Form1.frm".to_string(), src.as_bytes().to_vec())]);
    let refresh = src.find("Refresh").unwrap() as u32;
    let sites: Vec<_> = s.references(0, refresh, false).iter().map(|l| (l.span.start, text(src, l.span.start, l.span.len))).collect();
    assert_eq!(sites, vec![(src.find("Me.Refresh").unwrap() as u32 + 3, "Refresh")]);
    let helper = src.find("Helper").unwrap() as u32;
    assert_eq!(s.references(0, helper, false), vec![], "a Private procedure is not reachable through Me");
}

#[test]
fn me_member_references_a_public_variable_of_the_class() {
    let src = "Public Caption As String\nSub T()\n    sText = Me.Caption\nEnd Sub\n";
    let s = Session::from_sources(vec![("Model.cls".to_string(), src.as_bytes().to_vec())]);
    let decl = src.find("Caption").unwrap() as u32;
    let sites: Vec<_> = s.references(0, decl, false).iter().map(|l| text(src, l.span.start, l.span.len).to_string()).collect();
    assert_eq!(sites, vec!["Caption".to_string()]);
}
