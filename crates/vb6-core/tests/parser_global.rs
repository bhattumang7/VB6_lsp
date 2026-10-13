//! Module-level `Global` declares public variables and constants, exactly as
//! `Public` does.

use vb6_core::frontend::ast::{AstLit, ExprArena, ExprNode, NodeId, UnOpKind};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::session::Session;

struct Parsed {
    ctx: ScannerContext,
    nodes: Vec<NodeId>,
    arena: ExprArena,
    public: std::collections::HashMap<u32, bool>,
    diags: usize,
}

fn parse(src: &str) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().len();
    let public = std::mem::take(&mut parser.decl_public);
    drop(parser);
    Parsed { ctx, nodes, arena, public, diags }
}

fn value(p: &Parsed, id: NodeId) -> String {
    match p.arena.get(id) {
        ExprNode::Literal { lit: AstLit::Int(n) } => n.to_string(),
        ExprNode::Literal { lit: AstLit::Str(s) } => format!("{s:?}"),
        ExprNode::UnOp { op: UnOpKind::Neg, operand } => format!("-{}", value(p, *operand)),
        ExprNode::Paren { inner } => format!("({})", value(p, *inner)),
        _ => panic!("unexpected value"),
    }
}

/// Each top-level statement: [(is public, rendered DimItem)].
fn decls(p: &Parsed) -> Vec<Vec<(Option<bool>, String)>> {
    let name = |sym: u32| p.ctx.symbol(sym as usize).name.clone();
    p.nodes
        .iter()
        .map(|n| {
            let ExprNode::Block { stmts } = p.arena.get(*n) else { panic!("expected a declaration Block") };
            let items = stmts
                .iter()
                .map(|s| match p.arena.get(*s) {
                    ExprNode::DimItem { name: dn, is_const, type_node, init, .. } => (p.public.get(&s.0).copied(), format!(
                        "{} {} {} {}",
                        if *is_const { "const" } else { "dim" },
                        name(*dn),
                        match type_node.map(|t| p.arena.get(t)) {
                            Some(ExprNode::BuiltinType { kind }) => format!("type{kind}"),
                            Some(ExprNode::StringType { fixed_len: None }) => "String".into(),
                            None => "-".into(),
                            _ => panic!("unexpected type"),
                        },
                        init.map(|i| value(p, i)).unwrap_or_else(|| "-".into())
                    )),
                    _ => panic!("expected DimItem"),
                })
                .collect();
            items
        })
        .collect()
}

#[test]
fn global_const_list() {
    let p = parse("Global Const SQL_SUCCESS = 0\nGlobal Const SQL_NTS = -3, SQL_ERROR = (-1)\n");
    assert_eq!(p.diags, 0);
    assert_eq!(
        decls(&p),
        vec![
            vec![(Some(true), "const SQL_SUCCESS - 0".to_string())],
            vec![(Some(true), "const SQL_NTS - -3".to_string()), (Some(true), "const SQL_ERROR - (-1)".to_string())],
        ]
    );
}

#[test]
fn global_typed_const_and_variables() {
    let p = parse("Global Const DEFAULT_PROVIDER As String = \"SQLOLEDB\"\nGlobal g_nCount As Long, g_sName As String\n");
    assert_eq!(p.diags, 0);
    assert_eq!(
        decls(&p),
        vec![
            vec![(Some(true), "const DEFAULT_PROVIDER String \"SQLOLEDB\"".to_string())],
            vec![(Some(true), "dim g_nCount type3 -".to_string()), (Some(true), "dim g_sName String -".to_string())],
        ]
    );
}

#[test]
fn global_inside_conditional_compilation() {
    let p = parse("#If Win32 Then\nGlobal Const VK_TAB = 9\n#End If\n");
    assert_eq!(p.diags, 0);
}

#[test]
fn global_constant_resolves_from_another_module() {
    let consts = "Global Const SQL_NTS = -3\n";
    let user = "Sub T()\n    nLen = SQL_NTS\nEnd Sub\n";
    let s = Session::from_sources(vec![
        ("Consts.bas".to_string(), consts.as_bytes().to_vec()),
        ("User.bas".to_string(), user.as_bytes().to_vec()),
    ]);
    let at = (user.find("SQL_NTS").unwrap() + 1) as u32;
    let def = s.definition(1, at).expect("Global constant must resolve across modules");
    assert_eq!(def.module, 0);
    assert_eq!(&consts[def.span.start as usize..(def.span.start + def.span.len) as usize], "SQL_NTS");
}
