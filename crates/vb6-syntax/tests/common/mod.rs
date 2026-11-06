//! Shared helpers for parser AST tests.
#![allow(dead_code)]

pub mod answers;

use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprArena, ExprNode, NodeId, ProcKind, UnOpKind};
use vb6_syntax::frontend::declared_names::ModuleNames;
use vb6_syntax::frontend::parser::ModuleKind;
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

pub struct Parsed {
    pub ctx: ScannerContext,
    pub arena: ExprArena,
    pub top: Vec<NodeId>,
    pub diags: usize,
    /// Every diagnostic as `(code, start offset, length)`, in report order.
    pub found: Vec<(u32, u32, u32)>,
    /// Every diagnostic's message, in report order.
    pub messages: Vec<String>,
    /// The module-level names the declaration checks found.
    pub names: ModuleNames,
}

pub fn parse(src: &str) -> Parsed {
    parse_as(src, ModuleKind::Standard)
}

/// Parses `src` as a module of the given kind.
pub fn parse_kind(src: &str, kind: ModuleKind) -> Parsed {
    parse_as(src, kind)
}

/// Parses `src` as a module of the given kind.
pub fn parse_as(src: &str, kind: ModuleKind) -> Parsed {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let (top, diags, found, messages, names) = {
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
        (top, parser.diagnostics.items().len(), found, messages, parser.module_names.clone())
    };
    Parsed { ctx, arena, top, diags, found, messages, names }
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

    /// The tree below `id` as an S-expression, naming every part that matters.
    pub fn show(&self, id: NodeId) -> String {
        let list = |ids: &[NodeId]| ids.iter().map(|n| self.show(*n)).collect::<Vec<_>>().join(" ");
        let opt = |o: &Option<NodeId>| o.map_or("-".to_string(), |n| self.show(n));
        match self.node(id) {
            ExprNode::Literal { lit } => match lit {
                AstLit::Int(v) => format!("{v}%"),
                AstLit::Long(v) => format!("{v}&"),
                AstLit::Single(v) => format!("{v}!"),
                AstLit::Double(v) => format!("{v}#"),
                AstLit::Currency(v) => format!("{v}@"),
                AstLit::Str(s) => format!("{s:?}"),
                AstLit::Date(v) => format!("date({v})"),
                AstLit::Bool(b) => format!("{b}"),
                AstLit::Empty => "Empty".into(),
                AstLit::Null => "Null".into(),
            },
            ExprNode::NameRef { sym, .. } => self.name(*sym).to_string(),
            ExprNode::Nothing => "Nothing".into(),
            ExprNode::Me => "Me".into(),
            ExprNode::WithContext => "<with>".into(),
            ExprNode::Paren { inner } => format!("(paren {})", self.show(*inner)),
            ExprNode::BinOp { op, lhs, rhs } => format!("({} {} {})", bin_name(*op), self.show(*lhs), self.show(*rhs)),
            ExprNode::UnOp { op, operand } => {
                let n = match op {
                    UnOpKind::Neg => "neg",
                    UnOpKind::Pos => "pos",
                    UnOpKind::Not => "not",
                };
                format!("({n} {})", self.show(*operand))
            }
            ExprNode::MemberAccess { base, member, bang } => {
                format!("({} {} {})", if *bang { "bang" } else { "dot" }, self.show(*base), self.name(*member))
            }
            ExprNode::Call { func, args } => format!("(call {} {})", self.show(*func), self.show(*args)),
            ExprNode::ArgList { args } => format!("[{}]", list(args)),
            ExprNode::NamedArg { name, value } => format!("(named {} {})", self.name(*name), self.show(*value)),
            ExprNode::MissingArg => "<missing>".into(),
            ExprNode::CallStmt { callee, args } => format!("(Call {} {})", self.show(*callee), self.show(*args)),
            ExprNode::RaiseEvent { name, qualifier, args } => {
                format!("(RaiseEvent {} {} {})", opt(qualifier), self.name(*name), self.show(*args))
            }
            ExprNode::Block { stmts } => format!("{{{}}}", list(stmts)),
            ExprNode::With { obj, body } => format!("(With {} {})", self.show(*obj), self.show(*body)),
            ExprNode::BuiltinType { kind } => format!("type{kind}"),
            ExprNode::StringType { fixed_len } => format!("(String {})", opt(fixed_len)),
            ExprNode::UserType { name, child } => format!("(usertype {} {})", self.name(*name), opt(child)),
            ExprNode::ArrayType { elem } => format!("(array {})", self.show(*elem)),
            ExprNode::RangeTo { lo, hi } => format!("(to {} {})", self.show(*lo), self.show(*hi)),
            ExprNode::DimItem { name, is_const, is_static, is_new, with_events, bounds, type_node, init } => {
                let mut flags = String::new();
                for (on, tag) in [(*is_const, "const"), (*is_static, "static"), (*is_new, "new"), (*with_events, "withevents")] {
                    if on {
                        flags.push_str(&format!(" {tag}"));
                    }
                }
                format!("(item {}{flags} bounds={} type={} init={})", self.name(*name), opt(bounds), opt(type_node), opt(init))
            }
            ExprNode::ParamDef { flags, name, type_node, default } => {
                format!("(param {} flags={flags:#x} type={} default={})", self.name(*name), opt(type_node), opt(default))
            }
            ExprNode::EnumDecl { name, members } => format!("(Enum {} {})", self.name(*name), list(members)),
            ExprNode::TypeDecl { name, members } => format!("(Type {} {})", self.name(*name), list(members)),
            ExprNode::EventDecl { name, params } => format!("(Event {} {})", self.name(*name), self.show(*params)),
            ExprNode::ProcDecl { kind, name, params, ret_type, body } => {
                let k = match kind {
                    ProcKind::Sub => "Sub",
                    ProcKind::Function => "Function",
                    ProcKind::PropGet => "PropGet",
                    ProcKind::PropLet => "PropLet",
                    ProcKind::PropSet => "PropSet",
                };
                format!("({k} {} params={} ret={} {})", self.name(*name), opt(params), opt(ret_type), self.show(*body))
            }
            ExprNode::DeclareDecl { kind, name, params, ret_type, .. } => {
                let k = if *kind == ProcKind::Sub { "Sub" } else { "Function" };
                format!("(Declare{k} {} params={} ret={})", self.name(*name), opt(params), opt(ret_type))
            }
            ExprNode::OptionCompare { mode } => format!("(OptionCompare {mode})"),
            ExprNode::OptionBase { value } => format!("(OptionBase {value})"),
            _ => "?".into(),
        }
    }
}

fn bin_name(op: BinOpKind) -> &'static str {
    match op {
        BinOpKind::Add => "+",
        BinOpKind::Sub => "-",
        BinOpKind::Mul => "*",
        BinOpKind::Div => "/",
        BinOpKind::IDiv => "\\",
        BinOpKind::Mod => "Mod",
        BinOpKind::Pow => "^",
        BinOpKind::Cat => "&",
        BinOpKind::Eq => "=",
        BinOpKind::Ne => "<>",
        BinOpKind::Lt => "<",
        BinOpKind::Gt => ">",
        BinOpKind::Le => "<=",
        BinOpKind::Ge => ">=",
        BinOpKind::Like => "Like",
        BinOpKind::Is => "Is",
        BinOpKind::And => "And",
        BinOpKind::Or => "Or",
        BinOpKind::Xor => "Xor",
        BinOpKind::Eqv => "Eqv",
        BinOpKind::Imp => "Imp",
        BinOpKind::Not => "Not",
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
