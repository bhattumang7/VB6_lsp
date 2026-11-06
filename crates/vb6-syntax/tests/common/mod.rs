//! Shared helpers for parser AST tests.
#![allow(dead_code)]

pub mod answers;

use vb6_syntax::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_syntax::frontend::parser::{ModuleKind, Parser};
use vb6_syntax::frontend::scanner::ScannerContext;

pub struct Parsed {
    pub ctx: ScannerContext,
    pub arena: ExprArena,
    pub top: Vec<NodeId>,
    pub diags: usize,
    /// Every diagnostic as `(code, start offset, length)`, in report order.
    pub found: Vec<(u32, u32, u32)>,
    /// The text of every diagnostic, in report order (`Expected: ...` carries its list).
    pub messages: Vec<String>,
}

pub fn parse(src: &str) -> Parsed {
    parse_as(src, ModuleKind::Standard)
}

/// Parses `src` as a module of the given kind.
pub fn parse_as(src: &str, kind: ModuleKind) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let (top, diags, found, messages) = {
        let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), kind);
        let top = parser.parse_module(&mut arena);
        let found = parser
            .diagnostics
            .items()
            .iter()
            .map(|d| (d.code, d.span.start, d.span.len))
            .collect();
        let messages = parser
            .diagnostics
            .items()
            .iter()
            .map(|d| d.message().unwrap_or_else(|| format!("(code {:#x})", d.code)))
            .collect();
        (top, parser.diagnostics.items().len(), found, messages)
    };
    Parsed { ctx, arena, top, diags, found, messages }
}

impl Parsed {
    pub fn node(&self, id: NodeId) -> &ExprNode {
        self.arena.get(id)
    }

    /// Statements of the body of the `index`-th top-level procedure.
    pub fn proc_body(&self, index: usize) -> Vec<NodeId> {
        let procs: Vec<NodeId> = self
            .top
            .iter()
            .copied()
            .filter(|n| matches!(self.node(*n), ExprNode::ProcDecl { .. }))
            .collect();
        match self.node(procs[index]) {
            ExprNode::ProcDecl { body, .. } => self.block(*body),
            _ => unreachable!(),
        }
    }

    pub fn block(&self, id: NodeId) -> Vec<NodeId> {
        match self.node(id) {
            ExprNode::Block { stmts } => stmts.clone(),
            other => panic!("expected Block, got {}", std::mem::discriminant(other) == std::mem::discriminant(&ExprNode::EndStmt)),
        }
    }

    pub fn name(&self, sym: u32) -> &str {
        self.ctx.sym_name(sym as usize)
    }
}

/// An expression or statement as one line of text, every part of it spelled out:
/// `(Add 1 (Pos 2))`, `(= i 1)`, `(callstmt f [a b])`.
pub fn show(p: &Parsed, id: NodeId) -> String {
    let list = |ids: &[NodeId]| ids.iter().map(|n| show(p, *n)).collect::<Vec<_>>().join(" ");
    match p.node(id) {
        ExprNode::Literal { lit } => format!("{lit:?}"),
        ExprNode::NameRef { sym, .. } => p.name(*sym).to_string(),
        ExprNode::Paren { inner } => format!("(paren {})", show(p, *inner)),
        ExprNode::BinOp { op, lhs, rhs } => format!("({op:?} {} {})", show(p, *lhs), show(p, *rhs)),
        ExprNode::UnOp { op, operand } => format!("({op:?} {})", show(p, *operand)),
        ExprNode::Assign { target, value } => format!("(= {} {})", show(p, *target), show(p, *value)),
        ExprNode::ArgList { args } => list(args),
        ExprNode::Call { func, args } => format!("(call {} [{}])", show(p, *func), show(p, *args)),
        ExprNode::CallStmt { callee, args } => format!("(callstmt {} [{}])", show(p, *callee), show(p, *args)),
        _ => panic!("show: a node kind this helper does not spell out"),
    }
}

/// The statements of the first procedure, spelled out with [`show`].
pub fn body_shapes(p: &Parsed) -> Vec<String> {
    p.proc_body(0).iter().map(|n| show(p, *n)).collect()
}

/// Byte offset of the `n`-th (0-based) occurrence of `needle` in `src`.
pub fn offset_of(src: &str, needle: &str, n: usize) -> u32 {
    src.match_indices(needle).nth(n).expect("needle not found").0 as u32
}
