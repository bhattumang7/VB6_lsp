//! The names a module declares, and the rule that a name may not be declared twice.
//!
//! Every name declared at module level goes in one table: procedures, properties, `Declare`d
//! procedures, variables, constants, `WithEvents` variables and `Enum` members. Types and
//! enumerations have a table of their own, and so do events. Two declarations of a name clash
//! ("Ambiguous name detected") except where the language allows the pair: `Property Get`,
//! `Property Let` and `Property Set` of one name, and the same member name in two enumerations.
//! Letter case and a type-declaration character do not tell names apart.
//!
//! Only the first clash is reported. The declarations are checked in a fixed order, not in the
//! order they were written, and the clash is reported on the declaration being checked when it
//! meets the earlier one:
//!
//! 1. procedures, properties and `Declare`d procedures, in source order;
//! 2. the members of every `Enum`, in source order;
//! 3. the names of types and enumerations, in source order, each followed by a check that none
//!    of its own members is repeated (reported on the first of the two);
//! 4. `WithEvents` variables;
//! 5. variables and constants;
//! 6. events.
//!
//! The name in the message is spelled as the last declaration of the name in the module spells
//! it.

use std::collections::HashMap;

use crate::frontend::ast::{ExprArena, ExprNode, NodeId, NodeSpans, ProcKind};
use crate::frontend::diagnostics::{Diagnostic, ERR_AMBIGUOUS_NAME};
use crate::frontend::token::{Kw, Span};

/// The type a name is declared with.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DeclaredType {
    Boolean,
    Byte,
    Integer,
    Long,
    Single,
    Double,
    Currency,
    Date,
    String,
    /// `String * n`.
    FixedString,
    Variant,
    Object,
    /// A class or user-defined type, by its name.
    Named(u32),
}

/// What a module-level name declares.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NameKind {
    Sub,
    Function,
    /// A `Declare`d Sub.
    DeclareSub,
    /// A `Declare`d Function.
    DeclareFunction,
    PropertyGet,
    PropertyLet,
    PropertySet,
    EnumMember,
    WithEvents,
    Variable,
    Constant,
}

impl NameKind {
    /// Whether two declarations of one name may stand together.
    fn allows(self, other: NameKind) -> bool {
        use NameKind::*;
        matches!(
            (self, other),
            (PropertyGet, PropertyLet)
                | (PropertyGet, PropertySet)
                | (PropertyLet, PropertyGet)
                | (PropertyLet, PropertySet)
                | (PropertySet, PropertyGet)
                | (PropertySet, PropertyLet)
                | (EnumMember, EnumMember)
        )
    }
}

/// A module-level declaration of a name.
#[derive(Debug, Clone)]
pub struct NameEntry {
    pub sym: u32,
    pub kind: NameKind,
    /// The declaring node: a `DimItem`, an `EnumDecl` member, a `ProcDecl` or a `DeclareDecl`.
    pub node: NodeId,
    pub span: Span,
    /// The declared type of a variable or constant; for a procedure, the type it returns.
    pub declared_type: Option<DeclaredType>,
    /// For an enumeration member: the index of its enumeration among the module's.
    pub enum_index: Option<usize>,
}

/// An enumeration or user-defined type declared in the module.
#[derive(Debug, Clone)]
pub struct TypeEntry {
    pub sym: u32,
    pub node: NodeId,
    pub span: Span,
    pub is_enum: bool,
    /// The declaring node's members, in order.
    pub members: Vec<NodeId>,
}

/// An event declared in the module.
#[derive(Debug, Clone)]
pub struct EventEntry {
    pub sym: u32,
    pub span: Span,
}

/// Everything a module declares at module level.
#[derive(Debug, Default, Clone)]
pub struct ModuleNames {
    pub entries: Vec<NameEntry>,
    by_sym: HashMap<u32, Vec<usize>>,
    pub types: Vec<TypeEntry>,
    pub events: Vec<EventEntry>,
    /// The module's own name, from `Attribute VB_Name`.
    pub module_name: Option<String>,
}

/// What a module's declarations need from the parser to be read.
pub struct NameSource<'a> {
    pub arena: &'a ExprArena,
    pub spans: &'a NodeSpans,
    pub src: &'a [u8],
    pub top: &'a [NodeId],
    /// The spelling of a name as first written.
    pub spelling: &'a dyn Fn(u32) -> String,
}

impl ModuleNames {
    /// Reads the module-level declarations of `top`.
    pub fn collect(source: &NameSource<'_>) -> ModuleNames {
        let mut names = ModuleNames::default();
        let deftypes = DefTypes::collect(source);
        for &node in source.top {
            match source.arena.get(node) {
                ExprNode::ProcDecl { kind, name, ret_type, .. } => {
                    let kind = match kind {
                        ProcKind::Sub => NameKind::Sub,
                        ProcKind::Function => NameKind::Function,
                        ProcKind::PropGet => NameKind::PropertyGet,
                        ProcKind::PropLet => NameKind::PropertyLet,
                        ProcKind::PropSet => NameKind::PropertySet,
                    };
                    let declared = ret_type.and_then(|t| declared_type_of(source.arena, Some(t)));
                    names.add(*name, kind, node, source.spans.get(node), declared, None);
                }
                ExprNode::DeclareDecl { kind, name, ret_type, .. } => {
                    let k = if *kind == ProcKind::Sub { NameKind::DeclareSub } else { NameKind::DeclareFunction };
                    let declared = ret_type.and_then(|t| declared_type_of(source.arena, Some(t)));
                    names.add(*name, k, node, source.spans.get(node), declared, None);
                }
                ExprNode::Block { stmts } => {
                    for &item in stmts {
                        if let ExprNode::DimItem { name, is_const, with_events, type_node, .. } = source.arena.get(item) {
                            let kind = if *is_const {
                                NameKind::Constant
                            } else if *with_events {
                                NameKind::WithEvents
                            } else {
                                NameKind::Variable
                            };
                            // A variable declared without a type takes the type its first letter
                            // has under the `Def<Type>` statements in force; a constant takes
                            // the type of its value.
                            let declared = match declared_type_of(source.arena, *type_node) {
                                Some(t) => Some(t),
                                None if kind == NameKind::Constant => None,
                                None => Some(deftypes.type_for(&(source.spelling)(*name))),
                            };
                            names.add(*name, kind, item, source.spans.get(item), declared, None);
                        }
                    }
                }
                ExprNode::EnumDecl { name, members } => {
                    let index = names.types.iter().filter(|t| t.is_enum).count();
                    names.types.push(TypeEntry {
                        sym: *name,
                        node,
                        span: source.spans.get(node),
                        is_enum: true,
                        members: members.clone(),
                    });
                    for &m in members {
                        if let ExprNode::DimItem { name, .. } = source.arena.get(m) {
                            names.add(*name, NameKind::EnumMember, m, source.spans.get(m), Some(DeclaredType::Long), Some(index));
                        }
                    }
                }
                ExprNode::TypeDecl { name, members } => {
                    names.types.push(TypeEntry {
                        sym: *name,
                        node,
                        span: source.spans.get(node),
                        is_enum: false,
                        members: members.clone(),
                    });
                }
                ExprNode::EventDecl { name, .. } => {
                    names.events.push(EventEntry { sym: *name, span: source.spans.get(node) });
                }
                ExprNode::Attribute { target: None, name, values } => {
                    let is_name = (source.spelling)(*name).eq_ignore_ascii_case("VB_Name");
                    if let (true, Some(&v)) = (is_name, values.first()) {
                        if let ExprNode::Literal { lit: crate::frontend::ast::AstLit::Str(s) } = source.arena.get(v) {
                            names.module_name = Some(s.to_lowercase());
                        }
                    }
                }
                _ => {}
            }
        }
        names
    }

    fn add(
        &mut self,
        sym: u32,
        kind: NameKind,
        node: NodeId,
        span: Span,
        declared_type: Option<DeclaredType>,
        enum_index: Option<usize>,
    ) {
        let index = self.entries.len();
        self.entries.push(NameEntry { sym, kind, node, span, declared_type, enum_index });
        self.by_sym.entry(sym).or_default().push(index);
    }

    /// The declarations of a name, in the order they were written.
    pub fn declarations_of(&self, sym: u32) -> impl Iterator<Item = &NameEntry> {
        self.by_sym.get(&sym).into_iter().flatten().map(|&i| &self.entries[i])
    }

    /// The entry for a declaring node.
    pub fn entry_of(&self, node: NodeId) -> Option<&NameEntry> {
        self.entries.iter().find(|e| e.node == node)
    }

    /// The first clash of two declarations, if the module has one.
    pub fn first_clash(&self, source: &NameSource<'_>) -> Option<Diagnostic> {
        let spell = Spellings::collect(source);
        let diag = |sym: u32, span: Span| Diagnostic::with_name(ERR_AMBIGUOUS_NAME, span, &spell.of(sym, source));

        // The table the first three passes fill, and the later ones probe.
        let mut seen: HashMap<u32, Vec<NameKind>> = HashMap::new();
        let clashes = |seen: &HashMap<u32, Vec<NameKind>>, sym: u32, kind: NameKind| {
            seen.get(&sym).is_some_and(|kinds| kinds.iter().any(|k| !k.allows(kind)))
        };
        let pass = |wanted: &dyn Fn(NameKind) -> bool, seen: &mut HashMap<u32, Vec<NameKind>>| {
            for e in self.entries.iter().filter(|e| wanted(e.kind)) {
                if clashes(seen, e.sym, e.kind) {
                    return Some(diag(e.sym, e.span));
                }
                seen.entry(e.sym).or_default().push(e.kind);
            }
            None
        };

        // 1. Procedures.
        let procs = |k: NameKind| {
            matches!(
                k,
                NameKind::Sub
                    | NameKind::Function
                    | NameKind::DeclareSub
                    | NameKind::DeclareFunction
                    | NameKind::PropertyGet
                    | NameKind::PropertyLet
                    | NameKind::PropertySet
            )
        };
        if let Some(d) = pass(&procs, &mut seen) {
            return Some(d);
        }
        // 2. Enumeration members.
        if let Some(d) = pass(&|k| k == NameKind::EnumMember, &mut seen) {
            return Some(d);
        }
        // 3. Type and enumeration names, each followed by a check of its own members.
        let mut type_names: Vec<u32> = Vec::new();
        for t in &self.types {
            if type_names.contains(&t.sym) {
                return Some(diag(t.sym, t.span));
            }
            type_names.push(t.sym);
            let member_syms: Vec<(u32, Span)> = t
                .members
                .iter()
                .filter_map(|&m| match source.arena.get(m) {
                    ExprNode::DimItem { name, .. } => Some((*name, source.spans.get(m))),
                    _ => None,
                })
                .collect();
            for (i, (sym, span)) in member_syms.iter().enumerate() {
                if member_syms[i + 1..].iter().any(|(s, _)| s == sym) {
                    return Some(diag(*sym, *span));
                }
            }
        }
        // 4. WithEvents variables, then 5. variables and constants.
        if let Some(d) = pass(&|k| k == NameKind::WithEvents, &mut seen) {
            return Some(d);
        }
        if let Some(d) = pass(&|k| matches!(k, NameKind::Variable | NameKind::Constant), &mut seen) {
            return Some(d);
        }
        // 6. Events.
        let mut events: Vec<u32> = Vec::new();
        for e in &self.events {
            if events.contains(&e.sym) {
                return Some(diag(e.sym, e.span));
            }
            events.push(e.sym);
        }
        None
    }
}

/// The declared type of a type specifier node.
pub fn declared_type_of(arena: &ExprArena, node: Option<NodeId>) -> Option<DeclaredType> {
    let node = node?;
    Some(match arena.get(node) {
        ExprNode::BuiltinType { kind } => match kind {
            2 => DeclaredType::Integer,
            3 => DeclaredType::Long,
            4 => DeclaredType::Single,
            5 => DeclaredType::Double,
            6 => DeclaredType::Currency,
            7 => DeclaredType::Date,
            9 => DeclaredType::Object,
            11 => DeclaredType::Boolean,
            12 => DeclaredType::Variant,
            17 => DeclaredType::Byte,
            _ => return None,
        },
        ExprNode::StringType { fixed_len: None } => DeclaredType::String,
        ExprNode::StringType { fixed_len: Some(_) } => DeclaredType::FixedString,
        ExprNode::UserType { name, .. } => DeclaredType::Named(*name),
        ExprNode::ArrayType { elem } => return declared_type_of(arena, Some(*elem)),
        _ => return None,
    })
}

/// The `Def<Type>` statements of a module: the type each first letter has.
struct DefTypes {
    by_letter: [Option<DeclaredType>; 26],
}

impl DefTypes {
    fn collect(source: &NameSource<'_>) -> DefTypes {
        let mut by_letter = [None; 26];
        for &node in source.top {
            let ExprNode::DefType { type_kw, ranges } = source.arena.get(node) else { continue };
            let Some(ty) = def_type_of(*type_kw) else { continue };
            for &(lo, hi) in ranges {
                let letter = |sym: u32| {
                    (source.spelling)(sym).bytes().next().filter(u8::is_ascii_alphabetic).map(|b| b.to_ascii_lowercase() - b'a')
                };
                let Some(a) = letter(lo) else { continue };
                let b = if hi == 0 { Some(a) } else { letter(hi) };
                let Some(b) = b else { continue };
                if a <= b {
                    for l in a..=b {
                        by_letter[l as usize] = Some(ty);
                    }
                }
            }
        }
        DefTypes { by_letter }
    }

    /// The type of a name declared without one: by its first letter, Variant when no
    /// statement names the letter or the name does not start with a letter.
    fn type_for(&self, name: &str) -> DeclaredType {
        name.bytes()
            .next()
            .filter(u8::is_ascii_alphabetic)
            .and_then(|b| self.by_letter[(b.to_ascii_lowercase() - b'a') as usize])
            .unwrap_or(DeclaredType::Variant)
    }
}

fn def_type_of(kw: u16) -> Option<DeclaredType> {
    [
        (Kw::DefBool, DeclaredType::Boolean),
        (Kw::DefByte, DeclaredType::Byte),
        (Kw::DefCur, DeclaredType::Currency),
        (Kw::DefDate, DeclaredType::Date),
        (Kw::DefDbl, DeclaredType::Double),
        (Kw::DefInt, DeclaredType::Integer),
        (Kw::DefLng, DeclaredType::Long),
        (Kw::DefObj, DeclaredType::Object),
        (Kw::DefSng, DeclaredType::Single),
        (Kw::DefStr, DeclaredType::String),
        (Kw::DefVar, DeclaredType::Variant),
    ]
    .iter()
    .find(|(k, _)| *k as u16 == kw)
    .map(|(_, t)| *t)
}

/// The spelling each name has in a message: as its last declaration writes it.
struct Spellings {
    last: HashMap<u32, (u32, Span)>,
}

impl Spellings {
    fn collect(source: &NameSource<'_>) -> Spellings {
        let mut last: HashMap<u32, (u32, Span)> = HashMap::new();
        for i in 0..source.arena.len() {
            let id = NodeId(i as u32);
            let sym = match source.arena.get(id) {
                ExprNode::DimItem { name, .. }
                | ExprNode::ParamDef { name, .. }
                | ExprNode::ProcDecl { name, .. }
                | ExprNode::DeclareDecl { name, .. }
                | ExprNode::EnumDecl { name, .. }
                | ExprNode::TypeDecl { name, .. }
                | ExprNode::EventDecl { name, .. } => *name,
                _ => continue,
            };
            let span = source.spans.get(id);
            if span.len == 0 {
                continue;
            }
            match last.get(&sym) {
                Some((_, prev)) if prev.start >= span.start => {}
                _ => {
                    last.insert(sym, (sym, span));
                }
            }
        }
        Spellings { last }
    }

    fn of(&self, sym: u32, source: &NameSource<'_>) -> String {
        match self.last.get(&sym) {
            Some((_, span)) => written_name(source.src, *span).unwrap_or_else(|| (source.spelling)(sym)),
            None => (source.spelling)(sym),
        }
    }
}

/// The name written at `span`, without brackets or a type-declaration character.
fn written_name(src: &[u8], span: Span) -> Option<String> {
    let bytes = src.get(span.start as usize..(span.start + span.len) as usize)?;
    let text = bytes.iter().map(|&b| b as char).collect::<String>();
    let text = text.trim_end_matches(['%', '&', '!', '#', '@', '$']);
    let text = text.strip_prefix('[').and_then(|t| t.strip_suffix(']')).unwrap_or(text);
    if text.is_empty() { None } else { Some(text.to_string()) }
}
