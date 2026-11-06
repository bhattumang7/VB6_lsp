//! Evaluation of constant expressions over the syntax tree.
//!
//! A constant expression is built from literals, the names of constants and of `Enum`
//! members, parentheses and the operators. In an `Enum` member, an array bound, the size of a
//! fixed-length string or the default of an `Optional` parameter it may also call a few
//! built-in functions (`Len`, `Int`, `CLng`, ...); the value of a `Const` may not call
//! anything. The tree is first walked to find what makes it not constant (a variable, a
//! procedure, a call that cannot be made at compile time), and only then is a value worked out,
//! so a structural error is reported before an error in a value.

use crate::frontend::ast::{AstLit, BinOpKind, ExprArena, ExprNode, NodeId, UnOpKind};
use crate::frontend::const_value::{
    call_function, is_constant_function, is_language_function, BinaryOp, CompareMode, ConstValue, EvalError, Operand,
};

/// What a name stands for, as far as the module being read tells.
#[derive(Debug, Clone, PartialEq)]
pub enum Binding {
    /// A constant whose value is known.
    Constant(Operand),
    /// A declared thing that is not a usable constant: a variable, an array, a function, a
    /// property, a type, or a constant or `Enum` member not yet evaluated.
    NotConstant,
    /// A Sub, which has no value to use.
    Sub,
    /// A name qualified by an enumeration or by the module itself that is not declared there.
    NoSuchMember,
    /// A name the module does not declare. Whether it is a constant (a project-wide or library
    /// constant) is decided where those are known.
    Unknown,
}

/// What a constant expression can ask about names.
pub trait ConstEnv {
    /// The binding of a bare name.
    fn lookup(&self, sym: u32) -> Binding;
    /// The binding of `base.member`, where `base` is a chain of names.
    fn lookup_qualified(&self, path: &[u32], member: u32) -> Binding;
    /// How strings compare.
    fn compare_mode(&self) -> CompareMode;
    /// The spelling of a name, for matching the built-in functions.
    fn spelling(&self, sym: u32) -> String;
}

/// Where the expression stands, which decides whether it may call functions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Context {
    /// The value of a `Const`.
    ConstValue,
    /// An `Enum` member's value, an array bound, a fixed-length string's size, an `Optional`
    /// parameter's default.
    Other,
}

/// The outcome of evaluating: a value, or the fact that it depends on a name this module does
/// not declare, so only the module that declares it can tell.
#[derive(Debug, Clone, PartialEq)]
pub enum Evaluated {
    Value(Operand),
    Unresolved,
}

/// Evaluates the expression at `node`.
pub fn evaluate(
    arena: &ExprArena,
    node: NodeId,
    env: &dyn ConstEnv,
    ctx: Context,
) -> Result<Evaluated, EvalError> {
    if let Some(err) = structure_error(arena, node, env, ctx) {
        return Err(err);
    }
    value_of(arena, node, env)
}

/// The first thing in the tree, in reading order, that cannot be part of a constant. A name or
/// call this module cannot tell about (it is not declared here and is no built-in of the
/// language) is not held against the expression.
fn structure_error(arena: &ExprArena, node: NodeId, env: &dyn ConstEnv, ctx: Context) -> Option<EvalError> {
    match arena.get(node) {
        ExprNode::Paren { inner } => structure_error(arena, *inner, env, ctx),
        ExprNode::UnOp { operand, .. } => structure_error(arena, *operand, env, ctx),
        ExprNode::BinOp { .. } => {
            // A long chain of operators leans to the left; walk it without recursing down it.
            let (leaf, spine) = left_spine(arena, node);
            structure_error(arena, leaf, env, ctx).or_else(|| {
                spine.iter().rev().find_map(|op| match arena.get(*op) {
                    ExprNode::BinOp { rhs, .. } => structure_error(arena, *rhs, env, ctx),
                    _ => None,
                })
            })
        }
        ExprNode::NameRef { sym, .. } => binding_error(env.lookup(*sym)),
        ExprNode::MemberAccess { base, member, bang, .. } => {
            let path = qualified_path(arena, *base)?;
            // A member of a constant has nothing to name it by.
            if let [first] = path.as_slice() {
                if matches!(env.lookup(*first), Binding::Constant(_)) {
                    return Some(if *bang { EvalError::QualifierMustBeCollection } else { EvalError::InvalidQualifier });
                }
            }
            if *bang {
                return None;
            }
            binding_error(env.lookup_qualified(&path, *member))
        }
        ExprNode::Call { func, args } => {
            let ExprNode::ArgList { args } = arena.get(*args) else { return None };
            let argument_error = || args.iter().find_map(|a| structure_error(arena, *a, env, ctx));
            let ExprNode::NameRef { sym, .. } = arena.get(*func) else {
                return structure_error(arena, *func, env, ctx).or_else(argument_error);
            };
            match env.lookup(*sym) {
                Binding::Unknown => {}
                // A constant takes no arguments.
                Binding::Constant(_) => return Some(EvalError::ExpectedArray),
                // A function or an array of the module: calling it is not constant.
                _ => return Some(EvalError::NotConstant),
            }
            let spelling = env.spelling(*sym);
            if is_language_function(&spelling) {
                let callable = ctx == Context::Other && is_constant_function(&spelling);
                return if callable { argument_error() } else { Some(EvalError::NotConstant) };
            }
            argument_error()
        }
        _ => None,
    }
}

fn binding_error(binding: Binding) -> Option<EvalError> {
    match binding {
        Binding::Constant(_) | Binding::Unknown => None,
        Binding::NotConstant => Some(EvalError::NotConstant),
        Binding::Sub => Some(EvalError::ExpectedFunctionOrVariable),
        Binding::NoSuchMember => Some(EvalError::MemberNotFound),
    }
}

/// Splits a chain of binary operators that leans to the left into its leftmost operand and the
/// operator nodes above it, outermost first.
fn left_spine(arena: &ExprArena, node: NodeId) -> (NodeId, Vec<NodeId>) {
    let mut spine = Vec::new();
    let mut cur = node;
    while let ExprNode::BinOp { lhs, .. } = arena.get(cur) {
        spine.push(cur);
        cur = *lhs;
    }
    (cur, spine)
}

/// The names of a qualifier written `a.b.c`, outermost first.
fn qualified_path(arena: &ExprArena, node: NodeId) -> Option<Vec<u32>> {
    match arena.get(node) {
        ExprNode::NameRef { sym, .. } => Some(vec![*sym]),
        ExprNode::MemberAccess { base, member, bang: false, .. } => {
            let mut path = qualified_path(arena, *base)?;
            path.push(*member);
            Some(path)
        }
        _ => None,
    }
}

fn value_of(arena: &ExprArena, node: NodeId, env: &dyn ConstEnv) -> Result<Evaluated, EvalError> {
    use Evaluated::{Unresolved, Value};
    match arena.get(node) {
        ExprNode::Literal { lit } => Ok(Value(Operand::plain(literal_value(lit)))),
        ExprNode::Nothing => Ok(Value(Operand::plain(ConstValue::Nothing))),
        ExprNode::Paren { inner } => value_of(arena, *inner, env),
        ExprNode::UnOp { op, operand } => match value_of(arena, *operand, env)? {
            Value(v) => Ok(Value(match op {
                UnOpKind::Neg => v.negate()?,
                UnOpKind::Pos => v.plus()?,
                UnOpKind::Not => v.not()?,
            })),
            Unresolved => Ok(Unresolved),
        },
        ExprNode::BinOp { .. } => {
            let (leaf, spine) = left_spine(arena, node);
            let mut acc = value_of(arena, leaf, env)?;
            for op_node in spine.iter().rev() {
                let ExprNode::BinOp { op, rhs, .. } = arena.get(*op_node) else { continue };
                let r = value_of(arena, *rhs, env)?;
                acc = match (acc, r, binary_op(*op)) {
                    (Value(l), Value(r), Some(op)) => Value(Operand::binary(op, &l, &r, env.compare_mode())?),
                    _ => Unresolved,
                };
            }
            Ok(acc)
        }
        ExprNode::NameRef { sym, .. } => match env.lookup(*sym) {
            Binding::Constant(v) => Ok(Value(v)),
            _ => Ok(Unresolved),
        },
        ExprNode::MemberAccess { base, member, .. } => {
            let Some(path) = qualified_path(arena, *base) else { return Ok(Unresolved) };
            match env.lookup_qualified(&path, *member) {
                Binding::Constant(v) => Ok(Value(v)),
                _ => Ok(Unresolved),
            }
        }
        ExprNode::Call { func, args } => {
            let ExprNode::NameRef { sym, .. } = arena.get(*func) else { return Ok(Unresolved) };
            let ExprNode::ArgList { args } = arena.get(*args) else { return Ok(Unresolved) };
            if !is_constant_function(&env.spelling(*sym)) {
                return Ok(Unresolved);
            }
            // One argument is evaluated; a call with another count is the call checker's.
            if args.len() != 1 {
                return Ok(Unresolved);
            }
            match value_of(arena, args[0], env)? {
                Value(v) => Ok(Value(call_function(&env.spelling(*sym), &v)?)),
                Unresolved => Ok(Unresolved),
            }
        }
        _ => Ok(Unresolved),
    }
}

fn literal_value(lit: &AstLit) -> ConstValue {
    match lit {
        AstLit::Int(v) => ConstValue::Integer(*v as i16),
        AstLit::Long(v) => ConstValue::Long(*v),
        AstLit::Single(v) => ConstValue::Single(*v),
        AstLit::Double(v) => ConstValue::Double(*v),
        AstLit::Currency(v) => ConstValue::Currency(*v),
        AstLit::Str(s) => ConstValue::Str(s.to_string()),
        AstLit::Date(v) => ConstValue::Date(*v),
        AstLit::Bool(b) => ConstValue::Boolean(*b),
        AstLit::Empty => ConstValue::Empty,
        AstLit::Null => ConstValue::Null,
    }
}

fn binary_op(op: BinOpKind) -> Option<BinaryOp> {
    Some(match op {
        BinOpKind::Add => BinaryOp::Add,
        BinOpKind::Sub => BinaryOp::Sub,
        BinOpKind::Mul => BinaryOp::Mul,
        BinOpKind::Div => BinaryOp::Div,
        BinOpKind::IDiv => BinaryOp::IDiv,
        BinOpKind::Mod => BinaryOp::Mod,
        BinOpKind::Pow => BinaryOp::Pow,
        BinOpKind::Cat => BinaryOp::Cat,
        BinOpKind::Eq => BinaryOp::Eq,
        BinOpKind::Ne => BinaryOp::Ne,
        BinOpKind::Lt => BinaryOp::Lt,
        BinOpKind::Gt => BinaryOp::Gt,
        BinOpKind::Le => BinaryOp::Le,
        BinOpKind::Ge => BinaryOp::Ge,
        BinOpKind::Like => BinaryOp::Like,
        BinOpKind::Is => BinaryOp::Is,
        BinOpKind::And => BinaryOp::And,
        BinOpKind::Or => BinaryOp::Or,
        BinOpKind::Xor => BinaryOp::Xor,
        BinOpKind::Eqv => BinaryOp::Eqv,
        BinOpKind::Imp => BinaryOp::Imp,
        // `a Not b` has no value as an operation on two operands.
        BinOpKind::Not => return None,
    })
}
