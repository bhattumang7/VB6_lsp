//! `Friend` declarations: public within the project (marked public in the
//! visibility table) and also recorded as Friend, so the project's exported
//! interface can leave them out. The binder keeps each parameter's declared
//! type name beside its procedure.

use vb6_core::frontend::ast::{ExprArena, ExprNode, NodeId, ProcKind};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::sema::bind;
use vb6_core::sema::symbol::TypeName;

struct Parsed {
    ctx: ScannerContext,
    arena: ExprArena,
    top: Vec<NodeId>,
    public: std::collections::HashMap<u32, bool>,
    friend: std::collections::HashSet<u32>,
    diags: usize,
}

fn parse(src: &str) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let top = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    let public = std::mem::take(&mut parser.decl_public);
    let friend = std::mem::take(&mut parser.decl_friend);
    drop(parser);
    Parsed { ctx, arena, top, public, friend, diags }
}

/// Each top-level procedure: (kind, name, is public, is Friend).
fn procs(p: &Parsed) -> Vec<(ProcKind, String, Option<bool>, bool)> {
    p.top
        .iter()
        .map(|&id| match p.arena.get(id) {
            ExprNode::ProcDecl { kind, name, .. } => (
                *kind,
                p.ctx.symbol(*name as usize).name.clone(),
                p.public.get(&id.0).copied(),
                p.friend.contains(&id.0),
            ),
            _ => panic!("expected ProcDecl"),
        })
        .collect()
}

#[test]
fn friend_procedures_are_public_and_friend() {
    let p = parse(
        "Friend Sub Prepare()\nEnd Sub\nFriend Function Count() As Long\nEnd Function\n\
         Friend Property Get Owner() As Object\nEnd Property\nFriend Property Let Size(ByVal n As Long)\nEnd Property\n",
    );
    assert_eq!(p.diags, 0);
    assert_eq!(
        procs(&p),
        vec![
            (ProcKind::Sub, "Prepare".into(), Some(true), true),
            (ProcKind::Function, "Count".into(), Some(true), true),
            (ProcKind::PropGet, "Owner".into(), Some(true), true),
            (ProcKind::PropLet, "Size".into(), Some(true), true),
        ]
    );
}

#[test]
fn other_access_modifiers_are_not_friend() {
    let p = parse("Public Sub A()\nEnd Sub\nPrivate Sub B()\nEnd Sub\nGlobal Sub C()\nEnd Sub\nSub D()\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    assert_eq!(
        procs(&p),
        vec![
            (ProcKind::Sub, "A".into(), Some(true), false),
            (ProcKind::Sub, "B".into(), Some(false), false),
            (ProcKind::Sub, "C".into(), Some(true), false),
            (ProcKind::Sub, "D".into(), None, false),
        ]
    );
    assert!(p.friend.is_empty());
}

#[test]
fn parameter_type_names_are_kept_per_procedure() {
    let p = parse(
        "Public Sub A(ByVal n As Long, rs As ADODB.Recordset, o As clsOrder, v)\nEnd Sub\n\
         Public Function B() As String\nEnd Function\nPublic Sub C(items() As clsItem)\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    let spans = vb6_core::frontend::ast::NodeSpans::new();
    let bound = bind(&p.ctx, &p.arena, &p.top, &spans, &p.public);
    let tn = |q: Option<&str>, n: &str| Some(TypeName { qualifier: q.map(String::from), name: n.into() });
    assert_eq!(
        bound.decl_types.proc_params,
        vec![
            vec![None, tn(Some("adodb"), "recordset"), tn(None, "clsorder"), None],
            vec![],
            vec![tn(None, "clsitem")],
        ]
    );
}
