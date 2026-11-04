//! `Static` after (or instead of) the access modifier makes a procedure's
//! locals static: `Private Static Sub`, `Friend Static Function`, …

use vb6_syntax::frontend::ast::{ExprArena, ExprNode, NodeId, ProcKind};
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

struct Out {
    arena: ExprArena,
    top: Vec<NodeId>,
    diags: usize,
    public: std::collections::HashMap<u32, bool>,
    friend: std::collections::HashSet<u32>,
    stat: std::collections::HashSet<u32>,
}

fn parse(src: &str) -> Out {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let top = parser.parse_module(&mut arena);
    Out {
        diags: parser.diagnostics.items().len(),
        public: std::mem::take(&mut parser.decl_public),
        friend: std::mem::take(&mut parser.decl_friend),
        stat: std::mem::take(&mut parser.decl_static),
        arena,
        top,
    }
}

fn procs(o: &Out) -> Vec<(NodeId, ProcKind)> {
    o.top
        .iter()
        .filter_map(|n| match o.arena.get(*n) {
            ExprNode::ProcDecl { kind, .. } => Some((*n, *kind)),
            _ => None,
        })
        .collect()
}

#[test]
fn private_static_sub_and_call() {
    let o = parse(
        "Private Static Sub CountCalls()\n    Dim n As Integer\n    n = n + 1\nEnd Sub\n\nPublic Sub UseCountCalls()\n    CountCalls\nEnd Sub\n",
    );
    assert_eq!(o.diags, 0);
    let p = procs(&o);
    assert_eq!(p.len(), 2);
    assert_eq!(o.public.get(&p[0].0 .0), Some(&false));
    assert!(o.stat.contains(&p[0].0 .0));
    assert!(!o.stat.contains(&p[1].0 .0));
    let ExprNode::ProcDecl { body, .. } = o.arena.get(p[0].0) else { panic!() };
    let ExprNode::Block { stmts } = o.arena.get(*body) else { panic!() };
    assert_eq!(stmts.len(), 2);
}

#[test]
fn every_static_procedure_spelling() {
    let o = parse(
        "Public Static Function F() As Long\nEnd Function\n\
         Friend Static Sub G()\nEnd Sub\n\
         Static Sub H()\nEnd Sub\n\
         Private Static Property Get P() As Long\nEnd Property\n\
         Public Sub Plain()\nEnd Sub\n",
    );
    assert_eq!(o.diags, 0);
    let p = procs(&o);
    let kinds: Vec<_> = p.iter().map(|(_, k)| *k).collect();
    assert_eq!(kinds, vec![ProcKind::Function, ProcKind::Sub, ProcKind::Sub, ProcKind::PropGet, ProcKind::Sub]);
    let st: Vec<bool> = p.iter().map(|(n, _)| o.stat.contains(&n.0)).collect();
    assert_eq!(st, vec![true, true, true, true, false]);
    assert_eq!(o.public.get(&p[0].0 .0), Some(&true));
    assert!(o.friend.contains(&p[1].0 .0));
    assert_eq!(o.public.get(&p[2].0 .0), None);
    assert_eq!(o.public.get(&p[3].0 .0), Some(&false));
}
