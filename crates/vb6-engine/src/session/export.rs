//! The type library a project compiles to, described as a [`Library`].
//!
//! Only what VB6 places in the type library is described: the public class
//! modules (by `VB_Exposed`) of an ActiveX project and the public designers
//! (UserControls, UserDocuments) of a project that can expose them, each with
//! its public members, events, implemented interfaces, and the public enums and
//! types it declares. A Standard EXE exposes nothing. `Private` and `Friend`
//! members are absent: `Friend` is visible inside the project only.
//!
//! Types are spelled as the source writes them; an unqualified name that
//! resolves to a type of a referenced library is qualified with that library,
//! so a consumer of this library finds it.

use super::library::{
    Instancing, Library, LibraryMember, LibraryMemberKind, LibraryParam, LibraryType,
    LibraryTypeKind, ProjectKind,
};
use super::{ModuleData, Session};
use crate::frontend::ast::{AstLit, BinOpKind, ExprNode, NodeId, ProcKind, UnOpKind};
use crate::sema::symbol::{BoundProc, ParamFlags, TypeName};
use crate::sema::types::VbaType;

/// Module-level facts carried by `Attribute` lines.
#[derive(Default)]
struct ModuleAttributes {
    exposed: Option<bool>,
    creatable: Option<bool>,
    global_namespace: Option<bool>,
    /// (lowercase member name, dispatch id) from `VB_UserMemId` / `VB_VarUserMemId`.
    dispids: Vec<(String, i32)>,
    /// (lowercase member name, flags) from `VB_MemberFlags` / `VB_VarMemberFlags`.
    member_flags: Vec<(String, u32)>,
}

impl Session {
    /// The type library this session's project compiles to when it is built as
    /// `kind`. Its name is the session's project name.
    pub fn project_library(&self, kind: ProjectKind) -> Library {
        let mut types = Vec::new();
        if kind != ProjectKind::Exe {
            for (mi, m) in self.modules.iter().enumerate() {
                self.export_module(mi, m, kind, &mut types);
            }
        }
        Library { name: self.config.project_name.clone().unwrap_or_default(), types }
    }

    fn export_module(&self, mi: usize, m: &ModuleData, kind: ProjectKind, out: &mut Vec<LibraryType>) {
        let ext = m.path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
        let attrs = module_attributes(m);
        let exposed = attrs.exposed == Some(true);
        let instancing = match ext.as_str() {
            "cls" if exposed => Some(class_instancing(&attrs, m.multi_use)),
            // A public UserControl exists only in a control project; a public
            // UserDocument only in an ActiveX DLL or EXE.
            "ctl" if exposed && kind == ProjectKind::Control => None,
            "dob" if exposed && matches!(kind, ProjectKind::OleDll | ProjectKind::OleExe) => None,
            _ => return,
        };
        let x = Exporter { session: self, module: mi, m, attrs: &attrs };
        let mut members = Vec::new();
        let mut interfaces = Vec::new();
        let mut records = Vec::new();
        for &id in &m.top {
            x.collect(id, &mut members, &mut interfaces, &mut records);
        }
        out.push(LibraryType {
            name: m.module_name.clone(),
            kind: LibraryTypeKind::Class,
            interfaces,
            is_global: instancing.map_or(false, |i| matches!(i, Instancing::GlobalMultiUse | Instancing::GlobalSingleUse)),
            instancing,
            members,
            ..Default::default()
        });
        out.extend(records);
        for e in m.raw_bound.enum_decls.iter().filter(|e| e.is_public) {
            out.push(LibraryType {
                name: x.declared(e.name_span, e.sym_id),
                kind: LibraryTypeKind::Enum,
                members: e
                    .members
                    .iter()
                    .map(|em| LibraryMember {
                        name: x.declared(em.name_span, em.sym_id),
                        kind: LibraryMemberKind::Constant,
                        return_type: Some("Long".into()),
                        value: Some(em.value),
                        ..Default::default()
                    })
                    .collect(),
                ..Default::default()
            });
        }
    }
}

/// Instancing from the class attributes and header: not creatable is
/// PublicNotCreatable; otherwise MultiUse unless the header says one copy per
/// instance, Global when the members are global names.
fn class_instancing(attrs: &ModuleAttributes, multi_use: Option<bool>) -> Instancing {
    if attrs.creatable == Some(false) {
        return Instancing::PublicNotCreatable;
    }
    match (attrs.global_namespace == Some(true), multi_use.unwrap_or(true)) {
        (false, true) => Instancing::MultiUse,
        (false, false) => Instancing::SingleUse,
        (true, true) => Instancing::GlobalMultiUse,
        (true, false) => Instancing::GlobalSingleUse,
    }
}

fn module_attributes(m: &ModuleData) -> ModuleAttributes {
    let mut a = ModuleAttributes::default();
    let name_of = |s: u32| m.ctx.symbol(s as usize).name.to_ascii_lowercase();
    // A member's attributes sit inside its procedure, so every node is looked at.
    for i in 0..m.arena.len() as u32 {
        let ExprNode::Attribute { target, name, values } = m.arena.get(NodeId(i)) else { continue };
        let lit = values.first().and_then(|v| match m.arena.get(*v) {
            ExprNode::Literal { lit } => Some(lit.clone()),
            ExprNode::UnOp { op: UnOpKind::Neg, operand } => match m.arena.get(*operand) {
                ExprNode::Literal { lit: AstLit::Int(n) } => Some(AstLit::Int(-n)),
                ExprNode::Literal { lit: AstLit::Long(n) } => Some(AstLit::Long(-n)),
                _ => None,
            },
            _ => None,
        });
        let attr = name_of(*name);
        match (target.map(name_of), attr.as_str(), lit) {
            (None, "vb_exposed", Some(AstLit::Bool(b))) => a.exposed = Some(b),
            (None, "vb_creatable", Some(AstLit::Bool(b))) => a.creatable = Some(b),
            (None, "vb_globalnamespace", Some(AstLit::Bool(b))) => a.global_namespace = Some(b),
            (Some(t), "vb_usermemid" | "vb_varusermemid", Some(AstLit::Int(n) | AstLit::Long(n))) => {
                a.dispids.push((t, n))
            }
            (Some(t), "vb_memberflags" | "vb_varmemberflags", Some(AstLit::Str(s))) => {
                if let Ok(f) = u32::from_str_radix(s.trim(), 16) {
                    a.member_flags.push((t, f));
                }
            }
            _ => {}
        }
    }
    a
}

struct Exporter<'a> {
    session: &'a Session,
    module: usize,
    m: &'a ModuleData,
    attrs: &'a ModuleAttributes,
}

impl Exporter<'_> {
    fn spell(&self, sym: u32) -> String {
        self.m.ctx.symbol(sym as usize).name.clone()
    }

    /// A declared name as its declaration writes it (the symbol table keeps one
    /// spelling per name for the whole module).
    fn declared(&self, span: crate::frontend::ast::Span, sym: u32) -> String {
        let text = self.m.source.get(span.start as usize..(span.start + span.len) as usize).unwrap_or(&[]);
        let text: String = text.iter().map(|&b| b as char).collect();
        if span.len > 0 && text.eq_ignore_ascii_case(&self.spell(sym)) {
            text
        } else {
            self.spell(sym)
        }
    }

    fn declared_at(&self, id: NodeId, sym: u32) -> String {
        self.declared(self.m.spans.get(id), sym)
    }

    fn is_public(&self, id: NodeId, default: bool) -> bool {
        self.m.visibility.get(&id.0).copied().unwrap_or(default) && !self.m.friend.contains(&id.0)
    }

    fn member(&self, name: String, kind: LibraryMemberKind) -> LibraryMember {
        let lower = name.to_ascii_lowercase();
        let dispid = self.attrs.dispids.iter().find(|(n, _)| *n == lower).map(|(_, d)| *d);
        LibraryMember {
            is_default: dispid == Some(0),
            dispid,
            member_flags: self.attrs.member_flags.iter().find(|(n, _)| *n == lower).map(|(_, f)| *f),
            name,
            kind,
            ..Default::default()
        }
    }

    fn collect(
        &self,
        id: NodeId,
        members: &mut Vec<LibraryMember>,
        interfaces: &mut Vec<String>,
        records: &mut Vec<LibraryType>,
    ) {
        let m = self.m;
        match m.arena.get(id) {
            ExprNode::Block { stmts } => {
                for &s in stmts {
                    self.collect(s, members, interfaces, records);
                }
            }
            ExprNode::ProcDecl { kind, name, params, ret_type, .. } => {
                if !self.is_public(id, true) {
                    return;
                }
                let span = m.spans.get(id);
                let Some((pi, proc)) = m.raw_bound.procs.iter().enumerate().find(|(_, p)| p.name_span == span) else {
                    return;
                };
                let member_kind = match kind {
                    ProcKind::Sub | ProcKind::Function => LibraryMemberKind::Method,
                    ProcKind::PropGet => LibraryMemberKind::PropertyGet,
                    ProcKind::PropLet => LibraryMemberKind::PropertyLet,
                    ProcKind::PropSet => LibraryMemberKind::PropertySet,
                };
                let return_type = match kind {
                    ProcKind::Function | ProcKind::PropGet => self.type_text(*ret_type, &proc.ret_type, true),
                    _ => None,
                };
                let params = self.params(*params, Some((pi, proc)));
                members.push(LibraryMember {
                    return_type,
                    param_count: Some(params.len()),
                    params,
                    ..self.member(self.declared(proc.name_span, *name), member_kind)
                });
            }
            ExprNode::DimItem { name, is_const: false, type_node, bounds, .. } => {
                if !self.is_public(id, false) {
                    return;
                }
                let vba = match m.raw_bound.module_vars.iter().find(|v| v.name_span == m.spans.get(id)) {
                    Some(v) => v.vba_type.clone(),
                    None => VbaType::Variant,
                };
                let mut ty = self.type_text(*type_node, &vba, false);
                if bounds.is_some() && type_node.is_some() {
                    ty = ty.map(|t| if t.ends_with("()") { t } else { format!("{t}()") });
                }
                members.push(LibraryMember { return_type: ty, ..self.member(self.declared_at(id, *name), LibraryMemberKind::Field) });
            }
            ExprNode::EventDecl { name, params } => {
                let params = self.params(Some(*params), None);
                members.push(LibraryMember {
                    param_count: Some(params.len()),
                    params,
                    ..self.member(self.declared_at(id, *name), LibraryMemberKind::Event)
                });
            }
            ExprNode::Implements { library, name } => {
                let tn = TypeName {
                    qualifier: library.map(|l| self.spell(l).to_ascii_lowercase()),
                    name: self.spell(*name).to_ascii_lowercase(),
                };
                let written = match library {
                    Some(l) => format!("{}.{}", self.spell(*l), self.spell(*name)),
                    None => self.spell(*name),
                };
                interfaces.push(self.qualify(&tn, written));
            }
            ExprNode::TypeDecl { name, members: fields } => {
                if !self.is_public(id, true) {
                    return;
                }
                let decl = m.raw_bound.type_decls.iter().find(|t| t.name_span == m.spans.get(id));
                let mut out = Vec::new();
                for (fi, &f) in fields.iter().enumerate() {
                    let ExprNode::DimItem { name: fname, type_node, bounds, .. } = m.arena.get(f) else { continue };
                    let vba = decl.and_then(|d| d.members.get(fi)).map(|x| x.vba_type.clone()).unwrap_or(VbaType::Variant);
                    let mut ty = self.type_text(*type_node, &vba, false);
                    if bounds.is_some() && type_node.is_some() {
                        ty = ty.map(|t| if t.ends_with("()") { t } else { format!("{t}()") });
                    }
                    out.push(LibraryMember {
                        name: self.declared_at(f, *fname),
                        kind: LibraryMemberKind::Field,
                        return_type: ty,
                        ..Default::default()
                    });
                }
                records.push(LibraryType {
                    name: self.declared_at(id, *name),
                    kind: LibraryTypeKind::Record,
                    members: out,
                    ..Default::default()
                });
            }
            _ => {}
        }
    }

    /// The parameters of a procedure (bound, for its types) or an event (from
    /// the declaration alone).
    fn params(&self, list: Option<NodeId>, bound: Option<(usize, &BoundProc)>) -> Vec<LibraryParam> {
        let Some(ExprNode::ArgList { args }) = list.map(|l| self.m.arena.get(l)) else {
            return Vec::new();
        };
        args.iter()
            .enumerate()
            .filter_map(|(i, &a)| {
                let ExprNode::ParamDef { flags, name, type_node, default } = self.m.arena.get(a) else {
                    return None;
                };
                let f = ParamFlags::from_bits(*flags);
                let vba = match bound.and_then(|(_, p)| p.params.get(i)) {
                    Some(bp) => bp.vba_type.clone(),
                    None => {
                        let t = self.declared_vba(*type_node);
                        if f.is_array { VbaType::Array(Box::new(t)) } else { t }
                    }
                };
                Some(LibraryParam {
                    name: self.declared_at(a, *name),
                    type_name: self.type_text(*type_node, &vba, true),
                    // ByRef is the default; a ParamArray is always by reference.
                    by_ref: !f.by_val,
                    optional: f.optional,
                    default: default.map(|d| self.expr_text(d)),
                    param_array: f.param_array,
                })
            })
            .collect()
    }

    /// The type a declaration's type node states, for a declaration the binder
    /// did not record (an event parameter).
    fn declared_vba(&self, node: Option<NodeId>) -> VbaType {
        let mut node = match node {
            Some(n) => n,
            None => return VbaType::Variant,
        };
        while let ExprNode::ArrayType { elem } = self.m.arena.get(node) {
            node = *elem;
        }
        match self.m.arena.get(node) {
            ExprNode::BuiltinType { kind } => VbaType::from_kind(*kind),
            ExprNode::StringType { .. } => VbaType::String,
            ExprNode::UserType { name, .. } => VbaType::UserDefined(*name),
            _ => VbaType::Variant,
        }
    }

    /// A declared type as text: a named type as written (qualified with the
    /// library it resolves to), a built-in by its name, `None` for an
    /// undeclared Variant. `typed_by_suffix` admits a type given by a type
    /// character instead of a type node (`Function Name$()`).
    fn type_text(&self, node: Option<NodeId>, vba: &VbaType, typed_by_suffix: bool) -> Option<String> {
        let Some(mut n) = node else {
            return match vba {
                VbaType::Variant => None,
                _ if typed_by_suffix => Some(super::hover::type_str(&self.m.ctx, vba)),
                _ => None,
            };
        };
        let mut arrays = 0;
        while let ExprNode::ArrayType { elem } = self.m.arena.get(n) {
            arrays += 1;
            n = *elem;
        }
        let mut v = vba;
        let mut vba_arrays = 0;
        while let VbaType::Array(inner) = v {
            vba_arrays += 1;
            v = inner;
        }
        let base = match self.m.arena.get(n) {
            ExprNode::UserType { name, child } => {
                let first = self.spell(*name);
                match child.map(|c| self.m.arena.get(c)) {
                    Some(ExprNode::NameRef { sym, .. }) => format!("{first}.{}", self.spell(*sym)),
                    _ => {
                        let tn = TypeName { qualifier: None, name: first.to_ascii_lowercase() };
                        self.qualify(&tn, first)
                    }
                }
            }
            _ => super::hover::type_str(&self.m.ctx, v),
        };
        Some(format!("{base}{}", "()".repeat(arrays.max(vba_arrays))))
    }

    /// `written`, qualified with its library when it names a library's type
    /// without a qualifier.
    fn qualify(&self, tn: &TypeName, written: String) -> String {
        if tn.qualifier.is_some() {
            return written;
        }
        let lib = self.session.model.as_ref().and_then(|model| model.library_of_type(self.module, tn));
        match lib.and_then(|l| self.session.config.libraries.get(l)) {
            Some(l) => format!("{}.{written}", l.name),
            None => written,
        }
    }

    /// A default-value expression as VB source text.
    fn expr_text(&self, id: NodeId) -> String {
        let m = self.m;
        match m.arena.get(id) {
            ExprNode::Literal { lit } => lit_text(lit),
            ExprNode::NameRef { sym, .. } => self.spell(*sym),
            ExprNode::MemberAccess { base, member, bang } => {
                format!("{}{}{}", self.expr_text(*base), if *bang { "!" } else { "." }, self.spell(*member))
            }
            ExprNode::Paren { inner } => format!("({})", self.expr_text(*inner)),
            ExprNode::UnOp { op, operand } => {
                let o = match op {
                    UnOpKind::Neg => "-",
                    UnOpKind::Pos => "+",
                    UnOpKind::Not => "Not ",
                };
                format!("{o}{}", self.expr_text(*operand))
            }
            ExprNode::BinOp { op, lhs, rhs } => {
                format!("{} {} {}", self.expr_text(*lhs), binop_text(*op), self.expr_text(*rhs))
            }
            _ => {
                let s = m.spans.get(id);
                String::from_utf8_lossy(m.source.get(s.start as usize..(s.start + s.len) as usize).unwrap_or(&[])).into_owned()
            }
        }
    }
}

fn binop_text(op: BinOpKind) -> &'static str {
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
        BinOpKind::Dot => ".",
        BinOpKind::Bang => "!",
    }
}

fn lit_text(lit: &AstLit) -> String {
    match lit {
        AstLit::Int(n) | AstLit::Long(n) => n.to_string(),
        AstLit::Single(f) => f.to_string(),
        AstLit::Double(f) => f.to_string(),
        AstLit::Currency(c) => {
            // Currency is a scaled integer: four decimal places.
            let sign = if *c < 0 { "-" } else { "" };
            let a = c.unsigned_abs();
            let frac = a % 10_000;
            if frac == 0 {
                format!("{sign}{}", a / 10_000)
            } else {
                format!("{sign}{}.{}", a / 10_000, format!("{frac:04}").trim_end_matches('0'))
            }
        }
        AstLit::Str(s) => format!("\"{}\"", s.replace('"', "\"\"")),
        AstLit::Date(d) => date_text(*d),
        AstLit::Bool(true) => "True".into(),
        AstLit::Bool(false) => "False".into(),
        AstLit::Empty => "Empty".into(),
        AstLit::Null => "Null".into(),
    }
}

/// An OLE date (days since 30 December 1899, the fraction the time of day) as
/// a VB date literal: `#m/d/yyyy#`, with ` h:mm:ss AM` when it has a time.
fn date_text(d: f64) -> String {
    let days = d.trunc() as i64;
    // The fraction counts from midnight of the day, also for dates before 1899.
    let secs = (d.fract().abs() * 86_400.0).round() as i64;
    // Civil date from a day number (days since 1970-01-01).
    let z = days - 25_569 + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = yoe + era * 400 + if month <= 2 { 1 } else { 0 };
    let date = format!("{month}/{day}/{year}");
    if secs == 0 {
        if days == 0 {
            // Day zero with no time is the bare time 0: midnight.
            return "#12:00:00 AM#".into();
        }
        return format!("#{date}#");
    }
    let (h, mi, sc) = (secs / 3600, secs / 60 % 60, secs % 60);
    let (h12, ampm) = match h {
        0 => (12, "AM"),
        1..=11 => (h, "AM"),
        12 => (12, "PM"),
        _ => (h - 12, "PM"),
    };
    let time = format!("{h12}:{mi:02}:{sc:02} {ampm}");
    if days == 0 {
        format!("#{time}#")
    } else {
        format!("#{date} {time}#")
    }
}
