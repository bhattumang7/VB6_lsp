//! Shared helpers for parser AST tests.
#![allow(dead_code)]

pub mod answers;

use vb6_syntax::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

pub struct Parsed {
    pub ctx: ScannerContext,
    pub arena: ExprArena,
    pub top: Vec<NodeId>,
    pub diags: usize,
    /// Every diagnostic as `(code, start offset, length)`, in report order.
    pub found: Vec<(u32, u32, u32)>,
}

pub fn parse(src: &str) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let (top, diags, found) = {
        let mut parser = Parser::new(&mut ctx, src.as_bytes());
        let top = parser.parse_module(&mut arena);
        let found = parser
            .diagnostics
            .items()
            .iter()
            .map(|d| (d.code, d.span.start, d.span.len))
            .collect();
        (top, parser.diagnostics.items().len(), found)
    };
    Parsed { ctx, arena, top, diags, found }
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

/// Byte offset of the `n`-th (0-based) occurrence of `needle` in `src`.
pub fn offset_of(src: &str, needle: &str, n: usize) -> u32 {
    src.match_indices(needle).nth(n).expect("needle not found").0 as u32
}
