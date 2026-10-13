//! A line continuation (`_` at the end of a physical line) is whitespace to the
//! grammar wherever it appears, with LF or CR LF line breaks.

use vb6_core::frontend::ast::{AstLit, ExprArena, ExprNode, NodeId, ProcKind};
use vb6_core::frontend::parser::Parser;
use vb6_core::frontend::scanner::{Scanner, ScannerContext};
use vb6_core::frontend::token::TokenKind;

struct Parsed {
    ctx: ScannerContext,
    nodes: Vec<NodeId>,
    arena: ExprArena,
    diags: Vec<(u32, u32, u32)>,
}

fn parse(src: &str) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let nodes = parser.parse_module(&mut arena);
    let diags = parser.diagnostics.items().iter().map(|d| (d.code, d.span.start, d.span.len)).collect();
    drop(parser);
    Parsed { ctx, nodes, arena, diags }
}

fn crlf(src: &str) -> String {
    src.replace('\n', "\r\n")
}

/// S-expression of a node, with names spelled out.
fn render(p: &Parsed, id: NodeId) -> String {
    let name = |sym: u32| p.ctx.symbol(sym as usize).name.clone();
    let opt = |n: &Option<NodeId>| n.map(|x| render(p, x)).unwrap_or_else(|| "-".into());
    let list = |n: NodeId| match p.arena.get(n) {
        ExprNode::ArgList { args } => args.iter().map(|a| render(p, *a)).collect::<Vec<_>>().join(" "),
        _ => panic!("expected ArgList"),
    };
    match p.arena.get(id) {
        ExprNode::ProcDecl { kind, name: n, params, ret_type, body } => {
            let k = match kind {
                ProcKind::Sub => "Sub",
                ProcKind::Function => "Function",
                _ => panic!("unexpected proc kind"),
            };
            let ps = params.map(list).unwrap_or_default();
            format!("({k} {} [{ps}] {} {})", name(*n), opt(ret_type), render(p, *body))
        }
        ExprNode::DeclareDecl { kind, name: n, lib, alias, params, ret_type } => {
            let k = if *kind == ProcKind::Function { "Function" } else { "Sub" };
            let ps = params.map(list).unwrap_or_default();
            format!("(Declare {k} {} {} {} [{ps}] {})", name(*n), render(p, *lib), opt(alias), opt(ret_type))
        }
        ExprNode::Block { stmts } => {
            format!("{{{}}}", stmts.iter().map(|s| render(p, *s)).collect::<Vec<_>>().join(" "))
        }
        ExprNode::ParamDef { flags, name: n, type_node, default } => {
            format!("(param {flags:#x} {} {} {})", name(*n), opt(type_node), opt(default))
        }
        ExprNode::DimItem { name: n, is_const, type_node, init, .. } => {
            format!("({} {} {} {})", if *is_const { "const" } else { "dim" }, name(*n), opt(type_node), opt(init))
        }
        ExprNode::Assign { target, value } => format!("(= {} {})", render(p, *target), render(p, *value)),
        ExprNode::Call { func, args } => format!("(call {} [{}])", render(p, *func), list(*args)),
        ExprNode::NameRef { sym, .. } => name(*sym),
        ExprNode::BuiltinType { kind } => format!("type{kind}"),
        ExprNode::StringType { fixed_len: None } => "String".into(),
        ExprNode::AnyType => "Any".into(),
        ExprNode::Literal { lit: AstLit::Str(s) } => format!("{s:?}"),
        ExprNode::Literal { lit } => format!("{lit:?}"),
        _ => panic!("unexpected node"),
    }
}

/// Every top-level node rendered, for LF and CR LF versions of `src`; both
/// must agree and parse without diagnostics.
fn tree(src: &str) -> String {
    let mut out = Vec::new();
    for s in [src.to_string(), crlf(src)] {
        let p = parse(&s);
        assert_eq!(p.diags, vec![], "diagnostics for {s:?}");
        out.push(p.nodes.iter().map(|n| render(&p, *n)).collect::<Vec<_>>().join("\n"));
    }
    assert_eq!(out[0], out[1], "LF and CR LF must parse alike");
    out.remove(0)
}

#[test]
fn procedure_header_split_before_paren_inside_list_and_before_as() {
    let src = "Public Function Span _\n    ( _\n        ByVal dtFrom As Date, _\n        ByVal nDays As Integer) _\n    As Long\nEnd Function\n";
    assert_eq!(tree(src), "(Function Span [(param 0x2 dtFrom type7 -) (param 0x2 nDays type2 -)] type3 {})");
}

#[test]
fn continuation_before_closing_paren_of_parameter_list() {
    let src = "Sub Apply(ByRef nCount As Long, _\n          ByRef sName As String _\n         )\nEnd Sub\n";
    assert_eq!(tree(src), "(Sub Apply [(param 0x4 nCount type3 -) (param 0x4 sName String -)] - {})");
}

#[test]
fn constant_split_before_as() {
    let src = "Public Const KEY_NAME _\n                As String = \"Agencies\"\n";
    assert_eq!(tree(src), "{(const KEY_NAME String \"Agencies\")}");
}

#[test]
fn declare_split_between_every_part() {
    let src = "Private Declare Function SendMsg _\n    Lib \"user32\" _\n    Alias \"SendMessageA\" _\n    (ByVal hWnd As Long, _\n     ByVal lParam As Any) _\n    As Long\n";
    assert_eq!(
        tree(src),
        "(Declare Function SendMsg \"user32\" \"SendMessageA\" [(param 0x2 hWnd type3 -) (param 0x2 lParam Any -)] type3)"
    );
}

#[test]
fn call_split_before_paren_and_between_arguments() {
    let src = "Sub T()\nnTotal = Sum _\n    (nFirst, _\n     nSecond _\n    )\nEnd Sub\n";
    assert_eq!(tree(src), "(Sub T [] - {(= nTotal (call Sum [nFirst nSecond]))})");
}

#[test]
fn several_continuations_in_a_row_are_one_logical_line() {
    let src = "Sub T()\nnTotal = _\n _\n    nFirst\nEnd Sub\n";
    assert_eq!(tree(src), "(Sub T [] - {(= nTotal nFirst)})");
}

#[test]
fn crlf_is_a_single_end_of_line_token() {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut sc = Scanner::new(&mut ctx, b"a\r\nb\nc\r\rd");
    let mut kinds = Vec::new();
    loop {
        let t = sc.next_token();
        if t.kind == TokenKind::Eof {
            break;
        }
        kinds.push((t.kind == TokenKind::Eol, t.span.start));
    }
    // a, CRLF (one), b, LF, c, CR, CR (a lone CR is its own break), d
    assert_eq!(
        kinds,
        vec![(false, 0), (true, 1), (false, 3), (true, 4), (false, 5), (true, 6), (true, 7), (false, 8)]
    );
}
