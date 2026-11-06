//! The checks a module's declarations get once the whole module is read: the value of every
//! `Const` and `Enum` member, the bounds of every array, the size of every fixed-length
//! string, the default of every `Optional` parameter, and the rule that a name is declared
//! once.
//!
//! Declarations are taken in the order they are written. A constant may use the constants and
//! enumeration members declared before it (a name declared later, a variable, a procedure, a
//! type, is not a constant), and a procedure's own constants see what the procedure declared
//! before them as well as the module's names. A name the module does not declare cannot be
//! told here (it may be a constant of the project or of a library): an expression using one is
//! checked for everything else, and its value is left open.

use std::collections::HashMap;

use crate::frontend::ast::{ExprArena, ExprNode, NodeId, NodeSpans};
use crate::frontend::const_eval::{evaluate, Binding, ConstEnv, Context, Evaluated};
use crate::frontend::const_value::{CompareMode, ConstValue, EvalError, Operand, TargetType};
use crate::frontend::declared_names::{declared_type_of, DeclaredType, ModuleNames, NameKind, NameSource};
use crate::frontend::diagnostics::{
    Diagnostic, ERR_INVALID_DATA_TYPE_FOR_CONSTANT, ERR_INVALID_FIXED_STRING_LENGTH, ERR_RANGE_HAS_NO_VALUES,
};
use crate::frontend::token::Span;

/// How a name the module does not declare is resolved by whoever knows the project and its
/// libraries: the names of a qualified reference (`a.b.c`) as written.
pub type ExternalResolver = dyn Fn(&[String]) -> Binding;

/// What the checks need from the parser.
pub struct CheckInput<'a> {
    pub arena: &'a ExprArena,
    pub spans: &'a NodeSpans,
    pub src: &'a [u8],
    pub top: &'a [NodeId],
    /// The spelling of a name as first written.
    pub spelling: &'a dyn Fn(u32) -> String,
    pub external: Option<&'a ExternalResolver>,
}

/// The result of the checks.
pub struct Checked {
    pub diagnostics: Vec<Diagnostic>,
    pub names: ModuleNames,
}

/// The state of a constant after it is evaluated.
#[derive(Clone)]
enum Known {
    Value(Operand),
    /// The value depends on a name only another module can tell, or the declaration is wrong.
    Open,
}

struct Scope<'a> {
    names: &'a ModuleNames,
    values: &'a HashMap<NodeId, Known>,
    locals: Option<&'a HashMap<u32, Binding>>,
    mode: CompareMode,
    spelling: &'a dyn Fn(u32) -> String,
    external: Option<&'a ExternalResolver>,
}

impl Scope<'_> {
    fn module_binding(&self, sym: u32) -> Option<Binding> {
        let entry = self.names.declarations_of(sym).next()?;
        Some(match entry.kind {
            NameKind::Constant | NameKind::EnumMember => match self.values.get(&entry.node) {
                Some(Known::Value(v)) => Binding::Constant(v.clone()),
                Some(Known::Open) => Binding::Unknown,
                None => Binding::NotConstant,
            },
            NameKind::Sub | NameKind::DeclareSub => Binding::Sub,
            _ => Binding::NotConstant,
        })
    }

    /// A type or enumeration declared in the module, named as a value.
    fn type_binding(&self, sym: u32) -> Option<Binding> {
        self.names.types.iter().any(|t| t.sym == sym).then_some(Binding::NotConstant)
    }

    fn outside(&self, names: Vec<String>) -> Binding {
        match self.external {
            Some(resolve) => resolve(&names),
            None => Binding::Unknown,
        }
    }
}

impl ConstEnv for Scope<'_> {
    fn lookup(&self, sym: u32) -> Binding {
        if let Some(b) = self.locals.and_then(|l| l.get(&sym)) {
            return b.clone();
        }
        if let Some(b) = self.module_binding(sym) {
            return b;
        }
        if let Some(b) = self.type_binding(sym) {
            return b;
        }
        self.outside(vec![(self.spelling)(sym)])
    }

    fn lookup_qualified(&self, path: &[u32], member: u32) -> Binding {
        if let [first] = path {
            let own = (self.spelling)(*first).to_lowercase();
            if self.names.module_name.as_deref() == Some(own.as_str()) {
                return self.module_binding(member).or_else(|| self.type_binding(member)).unwrap_or(Binding::NoSuchMember);
            }
            if let Some(index) = self.names.types.iter().filter(|t| t.is_enum).position(|t| t.sym == *first) {
                return match self.names.entries.iter().find(|e| e.enum_index == Some(index) && e.sym == member) {
                    Some(e) => match self.values.get(&e.node) {
                        Some(Known::Value(v)) => Binding::Constant(v.clone()),
                        Some(Known::Open) => Binding::Unknown,
                        None => Binding::NotConstant,
                    },
                    None => Binding::NoSuchMember,
                };
            }
            if self.locals.is_some_and(|l| l.contains_key(first)) || self.names.declarations_of(*first).next().is_some() {
                // A member of a variable, a procedure's result, a constant.
                return Binding::NotConstant;
            }
        }
        let mut names: Vec<String> = path.iter().map(|s| (self.spelling)(*s)).collect();
        names.push((self.spelling)(member));
        self.outside(names)
    }

    fn compare_mode(&self) -> CompareMode {
        self.mode
    }

    fn spelling(&self, sym: u32) -> String {
        (self.spelling)(sym)
    }
}

struct Checker<'a> {
    input: &'a CheckInput<'a>,
    names: ModuleNames,
    values: HashMap<NodeId, Known>,
    locals: Option<HashMap<u32, Binding>>,
    mode: CompareMode,
    option_base: i32,
    diagnostics: Vec<Diagnostic>,
    /// The names of the parameters of the procedure being checked.
    parameters: Vec<u32>,
}

/// Runs every check on the module whose top-level nodes are `input.top`.
pub fn check_declarations(input: &CheckInput<'_>) -> Checked {
    let source = NameSource { arena: input.arena, spans: input.spans, src: input.src, top: input.top, spelling: input.spelling };
    let names = ModuleNames::collect(&source);
    let mut mode = CompareMode::Binary;
    let mut option_base = 0;
    for &node in input.top {
        match input.arena.get(node) {
            ExprNode::OptionCompare { mode: 1 } => mode = CompareMode::Text,
            ExprNode::OptionBase { value } => option_base = i32::from(*value),
            _ => {}
        }
    }
    let mut checker = Checker { input, names, values: HashMap::new(), locals: None, mode, option_base, diagnostics: Vec::new(), parameters: Vec::new() };
    let mut headers = Vec::new();
    // Declarations first, in the order written; procedures after, as they come.
    for &node in input.top {
        match input.arena.get(node) {
            ExprNode::Block { stmts } => {
                for &item in stmts {
                    checker.check_item(item);
                }
            }
            ExprNode::EnumDecl { members, .. } => checker.check_enum(members),
            ExprNode::TypeDecl { members, .. } => {
                for &m in members {
                    checker.check_item(m);
                }
            }
            ExprNode::ProcDecl { .. } | ExprNode::DeclareDecl { .. } | ExprNode::EventDecl { .. } => headers.push(node),
            _ => {}
        }
    }
    let mut declaration_diagnostics = std::mem::take(&mut checker.diagnostics);
    if let Some(clash) = checker.names.first_clash(&source) {
        declaration_diagnostics.push(clash);
    }
    for node in headers {
        checker.check_procedure(node);
    }
    declaration_diagnostics.append(&mut checker.diagnostics);
    Checked { diagnostics: declaration_diagnostics, names: checker.names }
}

impl Checker<'_> {
    fn scope(&self) -> Scope<'_> {
        Scope {
            names: &self.names,
            values: &self.values,
            locals: self.locals.as_ref(),
            mode: self.mode,
            spelling: self.input.spelling,
            external: self.input.external,
        }
    }

    fn eval(&self, node: NodeId, ctx: Context) -> Result<Evaluated, EvalError> {
        evaluate(self.input.arena, node, &self.scope(), ctx)
    }

    fn report(&mut self, code: u32, span: Span) {
        self.diagnostics.push(Diagnostic::new(code, span));
    }

    /// A `Dim`, `Const` or `Static` item, or a member of a user type.
    fn check_item(&mut self, item: NodeId) {
        let Some(DimShape { name, is_const, bounds, type_node, init }) = dim_shape(self.input.arena, item) else {
            return;
        };
        let span = self.input.spans.get(item);
        let outcome = self.check_item_parts(is_const, bounds, type_node, init);
        match outcome {
            Err(code) => {
                self.report(code, span);
                if is_const {
                    self.values.insert(item, Known::Open);
                }
            }
            Ok(known) => {
                if is_const {
                    self.values.insert(item, known);
                }
            }
        }
        if let Some(locals) = self.locals.as_mut() {
            let binding = match self.values.get(&item) {
                Some(Known::Value(v)) if is_const => Binding::Constant(v.clone()),
                Some(Known::Open) if is_const => Binding::Unknown,
                _ => Binding::NotConstant,
            };
            locals.insert(name, binding);
        }
    }

    fn check_item_parts(
        &mut self,
        is_const: bool,
        bounds: Option<NodeId>,
        type_node: Option<NodeId>,
        init: Option<NodeId>,
    ) -> Result<Known, u32> {
        if let Some(t) = type_node {
            self.check_fixed_length(t)?;
        }
        if let Some(b) = bounds {
            self.check_bounds(b)?;
        }
        if !is_const {
            return Ok(Known::Open);
        }
        let Some(init) = init else { return Ok(Known::Open) };
        let declared = declared_type_of(self.input.arena, type_node);
        // An object cannot be a constant, whatever its value.
        if matches!(declared, Some(DeclaredType::Object)) {
            return Err(ERR_INVALID_DATA_TYPE_FOR_CONSTANT);
        }
        match self.eval(init, Context::ConstValue).map_err(EvalError::code)? {
            Evaluated::Unresolved => {
                if type_node.is_some() && matches!(declared, Some(DeclaredType::Object)) {
                    return Err(ERR_INVALID_DATA_TYPE_FOR_CONSTANT);
                }
                Ok(Known::Open)
            }
            Evaluated::Value(op) => self.constant_value(op, declared, type_node.is_some()).map(Known::Value),
        }
    }

    /// The value a constant has once converted to its declared type (to the type of its value
    /// when it has none).
    fn constant_value(&self, op: Operand, declared: Option<DeclaredType>, typed: bool) -> Result<Operand, u32> {
        let convert = |v: &ConstValue, t: TargetType| v.convert(t).map_err(EvalError::code);
        match declared {
            None if typed => Err(ERR_INVALID_DATA_TYPE_FOR_CONSTANT),
            None => match op.value {
                ConstValue::Nothing => Err(ERR_INVALID_DATA_TYPE_FOR_CONSTANT),
                value => convert(&value, TargetType::Variant).map(Operand::plain),
            },
            Some(DeclaredType::Variant) => Ok(Operand { value: convert(&op.value, TargetType::Variant)?, variant: true }),
            Some(DeclaredType::Object) | Some(DeclaredType::Named(_)) => Err(ERR_INVALID_DATA_TYPE_FOR_CONSTANT),
            Some(t) => {
                let target = match t {
                    DeclaredType::Boolean => TargetType::Boolean,
                    DeclaredType::Byte => TargetType::Byte,
                    DeclaredType::Integer => TargetType::Integer,
                    DeclaredType::Long => TargetType::Long,
                    DeclaredType::Single => TargetType::Single,
                    DeclaredType::Double => TargetType::Double,
                    DeclaredType::Currency => TargetType::Currency,
                    DeclaredType::Date => TargetType::Date,
                    _ => TargetType::String,
                };
                Ok(Operand::plain(op.convert(target).map_err(EvalError::code)?))
            }
        }
    }

    /// The size of a fixed-length string: a constant from 1 to 65526 after rounding.
    fn check_fixed_length(&mut self, type_node: NodeId) -> Result<(), u32> {
        let ExprNode::StringType { fixed_len: Some(len) } = self.input.arena.get(type_node) else { return Ok(()) };
        // A length written as a number is checked where it is read.
        if matches!(self.input.arena.get(*len), ExprNode::Literal { .. }) {
            return Ok(());
        }
        match self.eval(*len, Context::Other).map_err(EvalError::code)? {
            Evaluated::Unresolved => Ok(()),
            Evaluated::Value(op) => {
                match op.value.as_long() {
                    Ok(n) if (1..65527).contains(&n) => Ok(()),
                    Ok(_) | Err(EvalError::Overflow) => Err(ERR_INVALID_FIXED_STRING_LENGTH),
                    Err(e) => Err(e.code()),
                }
            }
        }
    }

    /// The bounds of an array: constants, with no dimension running from a higher to a lower
    /// index.
    fn check_bounds(&mut self, bounds: NodeId) -> Result<(), u32> {
        let ExprNode::ArgList { args } = self.input.arena.get(bounds) else { return Ok(()) };
        let args = args.clone();
        for dim in args {
            let (lo, hi) = match self.input.arena.get(dim) {
                ExprNode::RangeTo { lo, hi } => (Some(*lo), *hi),
                _ => (None, dim),
            };
            let low = match lo {
                Some(l) => self.bound(l)?,
                None => Some(self.option_base),
            };
            let high = self.bound(hi)?;
            if let (Some(l), Some(h)) = (low, high) {
                if h < l {
                    return Err(ERR_RANGE_HAS_NO_VALUES);
                }
            }
        }
        Ok(())
    }

    fn bound(&mut self, node: NodeId) -> Result<Option<i32>, u32> {
        match self.eval(node, Context::Other).map_err(EvalError::code)? {
            Evaluated::Unresolved => Ok(None),
            Evaluated::Value(op) => op.value.as_long().map(Some).map_err(EvalError::code),
        }
    }

    /// The members of an enumeration: each takes its value as a Long, or one more than the
    /// member before it.
    fn check_enum(&mut self, members: &[NodeId]) {
        let mut previous: Option<Known> = None;
        for &m in members {
            let Some(DimShape { init, .. }) = dim_shape(self.input.arena, m) else { continue };
            let span = self.input.spans.get(m);
            let known = match init {
                Some(init) => match self.eval(init, Context::Other) {
                    Err(e) => {
                        self.report(e.code(), span);
                        Known::Open
                    }
                    Ok(Evaluated::Unresolved) => Known::Open,
                    Ok(Evaluated::Value(op)) => match op.value.convert(TargetType::Long) {
                        Ok(v) => Known::Value(Operand::plain(v)),
                        Err(e) => {
                            self.report(e.code(), span);
                            Known::Open
                        }
                    },
                },
                None => match &previous {
                    None => Known::Value(Operand::plain(ConstValue::Long(0))),
                    // One more than the largest Long wraps round to the smallest.
                    Some(Known::Value(Operand { value: ConstValue::Long(p), .. })) => {
                        Known::Value(Operand::plain(ConstValue::Long(p.wrapping_add(1))))
                    }
                    Some(_) => Known::Open,
                },
            };
            self.values.insert(m, known.clone());
            previous = Some(known);
        }
    }

    /// A procedure, a `Declare` or an event: the defaults of its `Optional` parameters, and for
    /// a procedure the declarations in its body.
    fn check_procedure(&mut self, node: NodeId) {
        let (params, body) = match self.input.arena.get(node) {
            ExprNode::ProcDecl { params, body, .. } => (*params, Some(*body)),
            ExprNode::DeclareDecl { params, .. } => (*params, None),
            ExprNode::EventDecl { params, .. } => (Some(*params), None),
            _ => return,
        };
        let span = self.input.spans.get(node);
        // The defaults are read in the module's scope; the body adds the parameters to it.
        self.locals = None;
        self.parameters.clear();
        if let Some(params) = params {
            if let ExprNode::ArgList { args } = self.input.arena.get(params) {
                for &p in args.clone().iter() {
                    self.check_parameter(p, span);
                }
            }
        }
        if let Some(body) = body {
            self.locals = Some(self.parameters.drain(..).map(|p| (p, Binding::NotConstant)).collect());
            let mut items = Vec::new();
            collect_items(self.input.arena, body, &mut items);
            items.sort_by_key(|n| n.0);
            for item in items {
                match self.input.arena.get(item) {
                    ExprNode::ReDimItem { type_node: Some(t), .. } => {
                        // A redimensioned array may be of fixed-length strings.
                        if let Err(code) = self.check_fixed_length(*t) {
                            self.report(code, self.input.spans.get(item));
                        }
                    }
                    _ => self.check_item(item),
                }
            }
        }
        self.locals = None;
    }

    fn check_parameter(&mut self, param: NodeId, header: Span) {
        let ExprNode::ParamDef { flags, name, type_node, default } = self.input.arena.get(param) else { return };
        let (flags, name, type_node, default) = (*flags, *name, *type_node, *default);
        if let (true, Some(default)) = (flags & 0x01 != 0, default) {
            let declared = declared_type_of(self.input.arena, type_node);
            if let Some(code) = self.default_error(default, declared, flags & 0x08 != 0) {
                self.report(code, header);
            }
        }
        // The parameters are in scope in the body only, not in each other's defaults.
        self.parameters.push(name);
    }

    /// The error a parameter's default value raises, if it does: it must be a constant that
    /// converts to the parameter's type.
    fn default_error(&self, default: NodeId, declared: Option<DeclaredType>, is_array: bool) -> Option<u32> {
        let value = match self.eval(default, Context::Other) {
            Err(e) => return Some(e.code()),
            Ok(Evaluated::Unresolved) => return None,
            Ok(Evaluated::Value(op)) => op,
        };
        if is_array {
            return None;
        }
        let target = match declared {
            None | Some(DeclaredType::Variant) => return None,
            Some(DeclaredType::Boolean) => TargetType::Boolean,
            Some(DeclaredType::Byte) => TargetType::Byte,
            Some(DeclaredType::Integer) => TargetType::Integer,
            Some(DeclaredType::Long) => TargetType::Long,
            Some(DeclaredType::Single) => TargetType::Single,
            Some(DeclaredType::Double) => TargetType::Double,
            Some(DeclaredType::Currency) => TargetType::Currency,
            Some(DeclaredType::Date) => TargetType::Date,
            Some(DeclaredType::String) | Some(DeclaredType::FixedString) => TargetType::String,
            // An object parameter's default is for its type to judge.
            Some(DeclaredType::Object) | Some(DeclaredType::Named(_)) => return None,
        };
        value.convert(target).err().map(EvalError::code)
    }
}

/// The declaration items among the statements below `node`, with the nodes they stand in.
fn collect_items(arena: &ExprArena, node: NodeId, out: &mut Vec<NodeId>) {
    if matches!(arena.get(node), ExprNode::DimItem { .. } | ExprNode::ReDimItem { .. }) {
        out.push(node);
    }
    let mut kids = Vec::new();
    arena.get(node).for_each_child(&mut |c| kids.push(c));
    for k in kids {
        collect_items(arena, k, out);
    }
}

/// The fields of a `DimItem` that the checks read.
struct DimShape {
    name: u32,
    is_const: bool,
    bounds: Option<NodeId>,
    type_node: Option<NodeId>,
    init: Option<NodeId>,
}

fn dim_shape(arena: &ExprArena, node: NodeId) -> Option<DimShape> {
    match arena.get(node) {
        ExprNode::DimItem { name, is_const, bounds, type_node, init, .. } => {
            Some(DimShape { name: *name, is_const: *is_const, bounds: *bounds, type_node: *type_node, init: *init })
        }
        _ => None,
    }
}
