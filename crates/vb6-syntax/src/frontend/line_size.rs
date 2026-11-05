//! How many bytes of compiled code a source line takes, for the limit one logical line has.
//!
//! The compiler keeps the code of one logical line in a buffer of [`MAX_LINE_CODE_SIZE`] bytes
//! and rejects the line with "Statement too complex" when it does not fit. This module holds
//! everything about that size in one place: the limit, what each kind of statement costs on top
//! of an assignment, and the size of the expressions whose cost is known exactly.
//!
//! Only shapes whose size has been established are measured. For any other shape
//! [`expression_size`] returns `None` and the line is not checked, so a size is never
//! overstated and a valid line is never rejected.

use crate::frontend::ast::{AstLit, BinOpKind, ExprArena, ExprNode, NodeId, UnOpKind};

/// Most bytes of compiled code one logical line may hold.
pub const MAX_LINE_CODE_SIZE: usize = 0x3c00;

/// What the parts of a line cost, in bytes.
pub mod cost {
    /// The store of an assignment (and what a condition's own code costs).
    pub const ASSIGNMENT: usize = 6;
    /// Each physical line of the logical line.
    pub const PER_PHYSICAL_LINE: usize = 4;
    /// A binary operator, a sign change, `Not`, a bracket pair, and the separator after an
    /// assignment that another statement follows on the line.
    pub const OPERATOR: usize = 2;
    /// The separator between the items of an output list.
    pub const ITEM_SEPARATOR: usize = 2;
    /// `Debug.Print` over an assignment of its first item.
    pub const DEBUG_PRINT: usize = 2;
    /// `Print #n` and `Write #n` over an assignment of the first item, without the channel.
    pub const FILE_OUTPUT: usize = 2;
    /// `Spc(n)` and `Tab(n)` over `n`.
    pub const SPC_OR_TAB: usize = 2;
    /// The clauses of an `Open` of a path over an assignment of the path.
    pub const OPEN: usize = 8;
    /// An `Open` with a `Len =` clause whose path is a name, over an assignment of the length.
    pub const OPEN_WITH_LENGTH: usize = 10;
    /// A `For` with no `Step`: the code of the loop itself, besides its variable and bounds.
    pub const FOR: usize = 2;
    /// A local `Const name =` over an assignment of its value: the name.
    pub const CONST: usize = 4;
    /// `Randomize seed` over an assignment of the seed.
    pub const RANDOMIZE: usize = 2;
    /// The record left out of a `Put #n, , value`.
    pub const MISSING_RECORD: usize = 2;
    /// `ReDim a(n)` over an assignment of the bound: the array name and the statement's code.
    pub const REDIM: usize = 8;
    /// `Lock #n, from` and `Unlock #n, from` over an assignment of `from`, besides the channel.
    pub const LOCK: usize = 2;
    /// `Width #n, w` over an assignment of `w`, besides the channel.
    pub const WIDTH: usize = 4;
    /// A name or Integer literal operand.
    pub const NAME_OR_INTEGER: usize = 4;
}

/// The bytes of the expression `value` alone: its operands and operators, not the store or the
/// lines. `None` for a shape that has not been measured.
///
/// An operand costs: a name or Integer literal 4, a Long or Single literal 6, a Double, Currency
/// or Date literal 10, a string 4 plus its length rounded up to even. Operator kinds and
/// operand kinds may be mixed, as they cost nothing more.
pub fn expression_part(arena: &ExprArena, value: NodeId) -> Option<usize> {
    let mut total = 0;
    let mut stack = vec![value];
    while let Some(id) = stack.pop() {
        match arena.get(id) {
            ExprNode::BinOp { op, lhs, rhs } => {
                if matches!(op, BinOpKind::Not) {
                    return None;
                }
                total += cost::OPERATOR;
                stack.push(*lhs);
                stack.push(*rhs);
            }
            ExprNode::UnOp { op: UnOpKind::Neg | UnOpKind::Not, operand } => {
                total += cost::OPERATOR;
                stack.push(*operand);
            }
            ExprNode::Paren { inner } => {
                total += cost::OPERATOR;
                stack.push(*inner);
            }
            ExprNode::NameRef { .. } => total += cost::NAME_OR_INTEGER,
            ExprNode::Literal { lit } => match lit {
                AstLit::Int(_) => total += cost::NAME_OR_INTEGER,
                AstLit::Long(_) | AstLit::Single(_) => total += 6,
                AstLit::Double(_) | AstLit::Currency(_) | AstLit::Date(_) => total += 10,
                AstLit::Str(s) => total += 4 + s.chars().count().next_multiple_of(2),
                _ => return None,
            },
            _ => return None,
        }
    }
    Some(total)
}

/// The compiled size of a line that is one statement whose cost over an assignment's store is
/// `extra` and whose expression is `value`, on `physical_lines` physical lines.
pub fn expression_size(arena: &ExprArena, value: NodeId, extra: usize, physical_lines: usize) -> Option<usize> {
    Some(cost::ASSIGNMENT + extra + cost::PER_PHYSICAL_LINE * physical_lines + expression_part(arena, value)?)
}

/// The compiled size of `target = value`: `None` unless the target is a plain name.
pub fn assignment_size(arena: &ExprArena, target: NodeId, value: NodeId, physical_lines: usize) -> Option<usize> {
    if !matches!(arena.get(target), ExprNode::NameRef { .. }) {
        return None;
    }
    expression_size(arena, value, 0, physical_lines)
}
