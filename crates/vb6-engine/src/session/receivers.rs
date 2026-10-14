//! What decides the object a member access is made on, and what flows into it.
//!
//! When a member access is bound at run time (its base is `Variant`/`Object`),
//! the object it runs on is whatever was last stored in the declaration the
//! base names. These queries expose the three facts a host needs to follow that
//! flow across a project, all from the bound model:
//!
//! * [`member_receivers`](Session::member_receivers): for each member access,
//!   the name whose declaration holds the base (`o` in `o.Save`, `Item` in
//!   `c.Item(1).Save`, the `With` object for `.Save`).
//! * [`object_stores`](Session::object_stores): every `Set`, `Let` and
//!   `For Each` and the static type of what it stores.
//! * [`call_arguments`](Session::call_arguments): every argument of every call,
//!   with its position and static type, so a parameter's values can be read off
//!   its callers.
//!
//! Names are given as spans: [`declaration_at`](Session::declaration_at) on a
//! span's start gives the declaration, so the host joins the three by it.

use super::Session;
use crate::frontend::ast::{ExprNode, NodeId, Span};
use crate::sema::project::{ProjectScope, SemType};
use crate::sema::types::VbaType;

/// The static type of an expression, as far as the binder knows it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StaticType {
    /// A class, form or other module of the project, by its name.
    Class(String),
    /// A type of an external library, `Library.Type`.
    Library(String),
    /// `Variant`, `Object` or untyped: decided at run time.
    LateBound,
    /// The `Nothing` literal.
    Nothing,
    /// A value type, an array or a user-defined type, as written.
    Value(String),
    /// A named type the project does not have, or one of a library that is
    /// not loaded.
    Unknown(String),
    /// The expression did not resolve.
    Unresolved,
}

/// The base of one member access.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Receiver {
    /// Span of the member name (the span [`member_accesses`](Session::member_accesses) reports).
    pub member: Span,
    /// Span of the name whose declaration holds the base: the variable, the
    /// called or indexed procedure or property, or the base of the enclosing
    /// `With`. `None` when the base is not a name (`Me`, a literal, an
    /// operator, `New`).
    pub head: Option<Span>,
    /// For a base that is a call or member access of a member of another object (`d.Item(k)`, `d.Items`),
    /// the name holding that other object (`d`): where the value was taken from.
    pub root: Option<Span>,
    /// The base is a call or index of the head (`f(1)`, `c(1)`, `c.Item(1)`):
    /// what a function returns, or an element or default member of what the
    /// head holds, not the head's own value.
    pub through_call: bool,
    /// The static type of the base.
    pub base_type: StaticType,
}

/// How a value is stored into a target.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StoreKind {
    /// `Set target = value`
    Set,
    /// `target = value` (or `Let target = value`)
    Let,
    /// `For Each target In value`: the target holds the elements of `value`.
    ForEach,
}

/// One store into a named target.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Store {
    pub kind: StoreKind,
    /// Span of the name stored into (for `Set a.b = v`, the span of `b`).
    pub target: Span,
    /// Span of the stored value's head name, when it has one; for `For Each`,
    /// of the collection's.
    pub value_head: Option<Span>,
    /// The static type of the value. For `For Each`, of an element: the
    /// array's element type, or `Unresolved` when the collection is not an
    /// array (the elements of a collection are not typed).
    pub value_type: StaticType,
}

/// One argument of one call.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Argument {
    /// Span of the name of the procedure called.
    pub callee: Span,
    /// For a method call (`coll.Add x`), the name whose declaration holds the object the method is
    /// called on, as for [`Receiver::head`].
    pub receiver: Option<Span>,
    /// Position of the argument, from 0; `None` for a named argument.
    pub index: Option<usize>,
    /// The name given to a named argument.
    pub name: Option<String>,
    /// Span of the argument's head name, when it has one.
    pub value_head: Option<Span>,
    pub value_type: StaticType,
}

impl Session {
    /// The base of every member access of `module`, in source order.
    pub fn member_receivers(&self, module: usize) -> Vec<Receiver> {
        let mut out = Vec::new();
        let Some(m) = self.modules.get(module) else {
            return out;
        };
        let mut with: Vec<NodeId> = Vec::new();
        for &top in &m.top {
            self.walk(module, top, &mut with, &mut |node, with| {
                if let ExprNode::MemberAccess { base, .. } = m.arena.get(node) {
                    out.push(Receiver {
                        member: m.spans.get(node),
                        head: self.head_of(module, *base, with),
                        root: self.root_of(module, *base, with),
                        through_call: self.through_call(module, *base, with),
                        base_type: self.static_type(module, *base),
                    });
                }
            });
        }
        out.sort_by_key(|r| r.member.start);
        out
    }

    /// Every `Set`, `Let` and `For Each` of `module` whose target is a name or
    /// a member, in source order.
    pub fn object_stores(&self, module: usize) -> Vec<Store> {
        let mut out = Vec::new();
        let Some(m) = self.modules.get(module) else {
            return out;
        };
        let mut with: Vec<NodeId> = Vec::new();
        for &top in &m.top {
            self.walk(module, top, &mut with, &mut |node, with| {
                let (kind, target, value) = match m.arena.get(node) {
                    ExprNode::SetAssign { target, value } => (StoreKind::Set, *target, *value),
                    ExprNode::Assign { target, value } => (StoreKind::Let, *target, *value),
                    ExprNode::ForEach {
                        var, collection, ..
                    } => (StoreKind::ForEach, *var, *collection),
                    _ => return,
                };
                let Some(target) = self.name_span(module, target) else {
                    return;
                };
                let value_type = match (kind, self.sem_of(module, value)) {
                    (StoreKind::ForEach, Some(SemType::Array(inner))) => self.render(&inner),
                    (StoreKind::ForEach, _) => StaticType::Unresolved,
                    _ => self.static_type(module, value),
                };
                out.push(Store {
                    kind,
                    target,
                    value_head: self.head_of(module, value, with),
                    value_type,
                });
            });
        }
        out.sort_by_key(|s| s.target.start);
        out
    }

    /// Every argument of every call of `module` to a named procedure, in
    /// source order.
    pub fn call_arguments(&self, module: usize) -> Vec<Argument> {
        let mut out = Vec::new();
        let Some(m) = self.modules.get(module) else {
            return out;
        };
        let mut with: Vec<NodeId> = Vec::new();
        for &top in &m.top {
            self.walk(module, top, &mut with, &mut |node, with| {
                let (func, args) = match m.arena.get(node) {
                    ExprNode::Call { func, args } => (*func, *args),
                    ExprNode::CallStmt { callee, args } => (*callee, *args),
                    _ => return,
                };
                let Some(callee) = self.name_span(module, func) else {
                    return;
                };
                let receiver = match m.arena.get(func) {
                    ExprNode::MemberAccess { base, .. } => self.head_of(module, *base, with),
                    _ => None,
                };
                let ExprNode::ArgList { args } = m.arena.get(args) else {
                    return;
                };
                for (i, &a) in args.iter().enumerate() {
                    let (index, name, value) = match m.arena.get(a) {
                        ExprNode::NamedArg { name, value } => (
                            None,
                            Some(m.ctx.symbol(*name as usize).name.clone()),
                            *value,
                        ),
                        _ => (Some(i), None, a),
                    };
                    out.push(Argument {
                        callee,
                        receiver,
                        index,
                        name,
                        value_head: self.head_of(module, value, with),
                        value_type: self.static_type(module, value),
                    });
                }
            });
        }
        out.sort_by_key(|a| a.callee.start);
        out
    }

    /// Visit every node under `id`, keeping the stack of `With` objects in force.
    fn walk(
        &self,
        module: usize,
        id: NodeId,
        with: &mut Vec<NodeId>,
        visit: &mut dyn FnMut(NodeId, &[NodeId]),
    ) {
        let m = &self.modules[module];
        visit(id, with);
        let node = m.arena.get(id);
        if let ExprNode::ForRange {
            loop_var,
            range,
            step,
        } = node
        {
            for raw in [*loop_var, *range, *step] {
                if raw != 0 {
                    self.walk(module, NodeId(raw), with, visit);
                }
            }
            return;
        }
        if let ExprNode::With { obj, body } = node {
            self.walk(module, *obj, with, visit);
            with.push(*obj);
            self.walk(module, *body, with, visit);
            with.pop();
            return;
        }
        let mut kids: Vec<NodeId> = Vec::new();
        node.for_each_child(&mut |c| kids.push(c));
        for c in kids {
            self.walk(module, c, with, visit);
        }
    }

    /// The span of the name whose declaration holds the value of `id`.
    fn head_of(&self, module: usize, id: NodeId, with: &[NodeId]) -> Option<Span> {
        let m = &self.modules[module];
        match m.arena.get(id) {
            ExprNode::NameRef { .. } | ExprNode::MemberAccess { .. } => Some(m.spans.get(id)),
            ExprNode::Call { func, .. } => self.head_of(module, *func, with),
            ExprNode::Paren { inner } => self.head_of(module, *inner, with),
            ExprNode::WithContext => {
                let (&obj, outer) = with.split_last()?;
                self.head_of(module, obj, outer)
            }
            _ => None,
        }
    }

    /// The name holding the object whose member `id` reads: for `d.Item(k)` and `d.Items` it is `d`.
    fn root_of(&self, module: usize, id: NodeId, with: &[NodeId]) -> Option<Span> {
        let m = &self.modules[module];
        match m.arena.get(id) {
            ExprNode::Call { func, .. } => self.root_of(module, *func, with),
            ExprNode::MemberAccess { base, .. } => self.head_of(module, *base, with),
            ExprNode::Paren { inner } => self.root_of(module, *inner, with),
            ExprNode::WithContext => {
                let (&obj, outer) = with.split_last()?;
                self.root_of(module, obj, outer)
            }
            _ => None,
        }
    }

    /// Whether the value of `id` is a call or index of its head.
    fn through_call(&self, module: usize, id: NodeId, with: &[NodeId]) -> bool {
        let m = &self.modules[module];
        match m.arena.get(id) {
            ExprNode::Call { .. } => true,
            ExprNode::Paren { inner } => self.through_call(module, *inner, with),
            ExprNode::WithContext => match with.split_last() {
                Some((&obj, outer)) => self.through_call(module, obj, outer),
                None => false,
            },
            _ => false,
        }
    }

    /// The span of a store target or callee that is a name or a member.
    fn name_span(&self, module: usize, id: NodeId) -> Option<Span> {
        let m = &self.modules[module];
        match m.arena.get(id) {
            ExprNode::NameRef { .. } | ExprNode::MemberAccess { .. } => Some(m.spans.get(id)),
            _ => None,
        }
    }

    fn sem_of(&self, module: usize, id: NodeId) -> Option<SemType> {
        let b = &self.modules[module].bound;
        if let Some(s) = b.sem_types.get(&id.0) {
            return Some(s.clone());
        }
        match b.types.get(&id.0) {
            Some(VbaType::Variant | VbaType::Object | VbaType::Any) | None => {
                Some(SemType::LateBound)
            }
            Some(_) => None,
        }
    }

    fn static_type(&self, module: usize, id: NodeId) -> StaticType {
        let m = &self.modules[module];
        if matches!(m.arena.get(id), ExprNode::Nothing) {
            return StaticType::Nothing;
        }
        match self.sem_of(module, id) {
            Some(s) => self.render(&s),
            None => match m.bound.types.get(&id.0) {
                Some(t) => StaticType::Value(format!("{t:?}")),
                None => StaticType::Unresolved,
            },
        }
    }

    fn render(&self, s: &SemType) -> StaticType {
        let name = || {
            self.model
                .as_ref()
                .map(|p| p.type_name(s))
                .unwrap_or_default()
        };
        match s {
            SemType::Class { .. } => StaticType::Class(name()),
            SemType::Lib { .. } => StaticType::Library(name()),
            SemType::LateBound => StaticType::LateBound,
            SemType::Value(_) | SemType::Udt { .. } | SemType::Array(_) => {
                StaticType::Value(name())
            }
            SemType::Unknown(n) | SemType::LibraryNotLoaded(n) => StaticType::Unknown(n.clone()),
            SemType::Namespace(_) | SemType::Unresolved => StaticType::Unresolved,
        }
    }
}
