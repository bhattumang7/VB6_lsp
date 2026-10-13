//! The session's view of the whole project, offered to the binder as a
//! [`ProjectScope`]: every module's declarations (as lowercase names, so it
//! owns its data and carries no borrow of the modules) plus the external
//! libraries the host described.

use std::collections::{HashMap, HashSet};

use super::library::{Library, LibraryMemberKind, LibraryType, LibraryTypeKind};
use super::module_index::ModuleIndex;
use super::ModuleData;
use crate::frontend::ast::ProcKind;
use crate::sema::project::{
    Access, Arity, MemberHit, MemberLookup, ProjectScope, Qualifier, SemType,
};
use crate::sema::symbol::{ExternalDecl, NameResolution, TypeName};
use crate::sema::types::VbaType;

/// What a type name refers to.
enum TypeRef {
    Udt { module: usize, idx: usize },
    Enum { module: usize, idx: usize },
    Class { module: usize },
    Lib { lib: usize, ty: usize },
    LibEnum { lib: usize, ty: usize },
    Unknown(String),
    NotLoaded(String),
}

/// A declared type: the `VbaType` plus the named type behind it, if any.
#[derive(Clone, PartialEq, Eq)]
struct Declared {
    vba: VbaType,
    name: Option<TypeName>,
}

#[derive(PartialEq, Eq)]
struct ProcSummary {
    name: String,
    kind: ProcKind,
    is_public: bool,
    params: usize,
    ret: Declared,
}

#[derive(PartialEq, Eq)]
struct VarSummary {
    name: String,
    is_public: bool,
    ty: Declared,
}

#[derive(PartialEq, Eq)]
struct TypeSummary {
    name: String,
    members: Vec<(String, Declared)>,
}

#[derive(PartialEq, Eq)]
struct EnumSummary {
    name: String,
    is_public: bool,
    members: Vec<String>,
}

#[derive(PartialEq, Eq)]
pub(super) struct ControlSummary {
    pub name: String,
    pub control_type: String,
    pub is_array: bool,
}

#[derive(PartialEq, Eq)]
struct ModuleSummary {
    /// Lowercase module name.
    name: String,
    /// The module name as written.
    display: String,
    /// A `.bas` standard module (a namespace), as opposed to a class, form or
    /// other designer (a type).
    standard: bool,
    predeclared: bool,
    global_namespace: bool,
    default_member: Option<String>,
    /// The library type a designer module extends (`VB.Form`).
    base_type: Option<TypeName>,
    procs: Vec<ProcSummary>,
    vars: Vec<VarSummary>,
    types: Vec<TypeSummary>,
    enums: Vec<EnumSummary>,
    types_public: Vec<bool>,
    controls: Vec<ControlSummary>,
}

struct LibMember {
    name: String,
    kind: LibraryMemberKind,
    return_type: Option<String>,
    is_default: bool,
    param_count: Option<usize>,
}

struct LibType {
    kind: LibraryTypeKind,
    extensible: bool,
    default_interface: Option<String>,
    interfaces: Vec<String>,
    extends: Vec<String>,
    members: Vec<LibMember>,
}

enum LibGlobal {
    /// A global instance of a type (`App`).
    Object(usize),
    /// A member of a module-kind type or an enum constant.
    Member { ty: usize, member: usize },
}

struct LibSummary {
    name: String,
    /// Type names as the library spells them, aligned with `types`.
    type_names: Vec<String>,
    types: Vec<LibType>,
    type_by_name: HashMap<String, usize>,
    globals: HashMap<String, LibGlobal>,
}

pub(super) struct ProjectModel {
    modules: Vec<ModuleSummary>,
    module_by_name: HashMap<String, usize>,
    index: ModuleIndex,
    project_name: Option<String>,
    libraries: Vec<LibSummary>,
    library_by_name: HashMap<String, usize>,
}

fn lc(s: &str) -> String {
    s.to_ascii_lowercase()
}

/// The accessor kinds that serve a use needing `access`.
fn serves(kind: ProcKind, access: Access) -> bool {
    match access {
        Access::Get => matches!(kind, ProcKind::Sub | ProcKind::Function | ProcKind::PropGet),
        Access::Let => kind == ProcKind::PropLet,
        Access::Set => kind == ProcKind::PropSet,
    }
}

/// The parameter list of a library member that is a procedure.
fn lib_callable(m: &LibMember) -> Option<Arity> {
    match m.kind {
        LibraryMemberKind::Method
        | LibraryMemberKind::PropertyGet
        | LibraryMemberKind::PropertyLet
        | LibraryMemberKind::PropertySet => Some(match m.param_count {
            Some(n) => Arity::Known(n),
            None => Arity::Unknown,
        }),
        LibraryMemberKind::Field | LibraryMemberKind::Constant | LibraryMemberKind::Event => None,
    }
}

fn lib_serves(kind: LibraryMemberKind, access: Access) -> bool {
    match access {
        Access::Get => matches!(
            kind,
            LibraryMemberKind::Method
                | LibraryMemberKind::PropertyGet
                | LibraryMemberKind::Field
                | LibraryMemberKind::Constant
        ),
        Access::Let => matches!(kind, LibraryMemberKind::PropertyLet | LibraryMemberKind::Field),
        Access::Set => matches!(kind, LibraryMemberKind::PropertySet | LibraryMemberKind::Field),
    }
}

impl ProjectModel {
    /// The identities `(library, declaring type, member)` that a member named
    /// `member` of type `type_name` of library `library` can resolve to: the
    /// member as declared by the type itself or by any interface it exposes.
    pub(super) fn library_member_keys(
        &self,
        library: &str,
        type_name: &str,
        member: &str,
    ) -> Vec<(usize, usize, usize)> {
        let Some(&lib) = self.library_by_name.get(&lc(library)) else {
            return Vec::new();
        };
        let Some(&ty) = self.libraries[lib].type_by_name.get(&lc(type_name)) else {
            return Vec::new();
        };
        let member = lc(member);
        let mut out = Vec::new();
        for (l, t) in self.lib_chain(lib, ty) {
            for (mi, m) in self.libraries[l].types[t].members.iter().enumerate() {
                // An event is handled, not called: no use of the name reaches it.
                if m.name == member && m.kind != LibraryMemberKind::Event {
                    out.push((l, t, mi));
                }
            }
        }
        out
    }

    /// Whether module `idx`, as the model saw it, has the same interface as
    /// `module` now: the same declarations, in the same order, with the same
    /// visibility and types, and the same name, kind and controls. When it
    /// does, nothing another module's binding depends on has changed.
    pub(super) fn module_interface_is(&self, idx: usize, module: &ModuleData) -> bool {
        self.modules.get(idx).map_or(false, |s| *s == summarize(module))
    }

    pub(super) fn build(
        modules: &[ModuleData],
        project_name: Option<&str>,
        libraries: &[Library],
    ) -> Self {
        let summaries: Vec<ModuleSummary> = modules.iter().map(summarize).collect();
        let mut module_by_name = HashMap::new();
        for (i, m) in summaries.iter().enumerate() {
            if !m.name.is_empty() {
                module_by_name.entry(m.name.clone()).or_insert(i);
            }
        }
        let mut index = ModuleIndex::new();
        for (i, m) in modules.iter().enumerate() {
            let s = &summaries[i];
            // Procedures and variables of a class are reached through an
            // instance, not by bare name; only a standard module (or a class
            // marked `VB_GlobalNameSpace`) contributes them as project names.
            let members_global = s.standard || s.global_namespace;
            index.add_module(
                i,
                &m.bound,
                |sym| m.ctx.symbol(sym as usize).name.to_ascii_lowercase(),
                members_global,
            );
        }
        let libs: Vec<LibSummary> = libraries.iter().map(summarize_library).collect();
        let mut library_by_name = HashMap::new();
        for (i, l) in libs.iter().enumerate() {
            library_by_name.entry(l.name.clone()).or_insert(i);
        }
        ProjectModel {
            modules: summaries,
            module_by_name,
            index,
            project_name: project_name.map(lc),
            libraries: libs,
            library_by_name,
        }
    }

    // ── Declared types ──────────────────────────────────────────────────────

    /// The static type of a declaration made in `module`.
    fn declared_sem(&self, module: usize, d: &Declared) -> SemType {
        match &d.vba {
            VbaType::Array(inner) => SemType::Array(Box::new(self.declared_sem(
                module,
                &Declared { vba: (**inner).clone(), name: d.name.clone() },
            ))),
            VbaType::Variant | VbaType::Object | VbaType::Any => SemType::LateBound,
            VbaType::UserDefined(_) => match &d.name {
                Some(tn) => self.sem_of_type_name(module, tn),
                None => SemType::LateBound,
            },
            other => SemType::Value(other.clone()),
        }
    }

    fn is_project(&self, q: &str) -> bool {
        self.project_name.as_deref() == Some(q)
    }

    fn find_type_in(&self, module: usize, name: &str) -> Option<TypeRef> {
        let m = &self.modules[module];
        if let Some(idx) = m.types.iter().position(|t| t.name == name) {
            return Some(TypeRef::Udt { module, idx });
        }
        m.enums.iter().position(|e| e.name == name).map(|idx| TypeRef::Enum { module, idx })
    }

    fn find_public_type_in(&self, module: usize, name: &str) -> Option<TypeRef> {
        let m = &self.modules[module];
        if let Some(idx) = m.types.iter().position(|t| t.name == name) {
            return m.types_public[idx].then_some(TypeRef::Udt { module, idx });
        }
        let idx = m.enums.iter().position(|e| e.name == name && e.is_public)?;
        Some(TypeRef::Enum { module, idx })
    }

    fn class_module(&self, name: &str) -> Option<usize> {
        self.module_by_name.get(name).copied().filter(|&i| !self.modules[i].standard)
    }

    fn lib_type(&self, lib: usize, name: &str) -> Option<TypeRef> {
        let ty = *self.libraries[lib].type_by_name.get(name)?;
        Some(match self.libraries[lib].types[ty].kind {
            LibraryTypeKind::Enum => TypeRef::LibEnum { lib, ty },
            _ => TypeRef::Lib { lib, ty },
        })
    }

    /// The type a declared type name denotes, seen from module `from`.
    ///
    /// Unqualified, the search order is: the module's own types, the project's
    /// classes, public types of every module, then the libraries. Qualified by
    /// the project, a module or a library, only that scope is searched; a
    /// qualifier that is none of them names a library that is not loaded.
    fn type_ref(&self, from: usize, tn: &TypeName) -> TypeRef {
        let name = tn.name.as_str();
        let unknown = || TypeRef::Unknown(name.to_string());
        match tn.qualifier.as_deref() {
            None => {
                if let Some(t) = self.find_type_in(from, name) {
                    return t;
                }
                if let Some(m) = self.class_module(name) {
                    return TypeRef::Class { module: m };
                }
                for m in 0..self.modules.len() {
                    if let Some(t) = self.find_public_type_in(m, name) {
                        return t;
                    }
                }
                for l in 0..self.libraries.len() {
                    if let Some(t) = self.lib_type(l, name) {
                        return t;
                    }
                }
                unknown()
            }
            Some(q) => {
                if self.is_project(q) {
                    if let Some(m) = self.class_module(name) {
                        return TypeRef::Class { module: m };
                    }
                    for m in 0..self.modules.len() {
                        if let Some(t) = self.find_public_type_in(m, name) {
                            return t;
                        }
                    }
                    return unknown();
                }
                if let Some(&m) = self.module_by_name.get(q) {
                    return self.find_public_type_in(m, name).unwrap_or_else(unknown);
                }
                if let Some(&l) = self.library_by_name.get(q) {
                    return self.lib_type(l, name).unwrap_or_else(unknown);
                }
                TypeRef::NotLoaded(q.to_string())
            }
        }
    }

    /// The library whose type an unqualified type name written in `from`
    /// denotes, by VB6's search order; `None` when it denotes a type of the
    /// project or nothing.
    pub(super) fn library_of_type(&self, from: usize, tn: &TypeName) -> Option<usize> {
        match self.type_ref(from, tn) {
            TypeRef::Lib { lib, .. } | TypeRef::LibEnum { lib, .. } => Some(lib),
            _ => None,
        }
    }

    // ── Library member lookup ───────────────────────────────────────────────

    /// Resolve a library type name written in `lib` (same library, or
    /// `Library.Type`).
    fn lib_type_ref(&self, lib: usize, name: &str) -> Option<(usize, usize)> {
        let name = lc(name);
        if let Some((q, n)) = name.split_once('.') {
            let l = *self.library_by_name.get(q)?;
            return Some((l, *self.libraries[l].type_by_name.get(n)?));
        }
        Some((lib, *self.libraries[lib].type_by_name.get(&name)?))
    }

    /// Types whose members `ty` exposes, in lookup order: itself, then (for a
    /// coclass) its default and other interfaces, each followed by the
    /// interfaces it extends.
    fn lib_chain(&self, lib: usize, ty: usize) -> Vec<(usize, usize)> {
        let mut out = Vec::new();
        let mut seen: HashSet<(usize, usize)> = HashSet::new();
        let mut stack = vec![(lib, ty)];
        while let Some((l, t)) = stack.pop() {
            if !seen.insert((l, t)) {
                continue;
            }
            out.push((l, t));
            let ty = &self.libraries[l].types[t];
            // Pushed in reverse so the first-listed is visited first.
            let mut next: Vec<(usize, usize)> = Vec::new();
            for n in ty.default_interface.iter().chain(&ty.interfaces).chain(&ty.extends) {
                if let Some(r) = self.lib_type_ref(l, n) {
                    next.push(r);
                }
            }
            stack.extend(next.into_iter().rev());
        }
        out
    }

    fn lib_return_sem(&self, lib: usize, ret: &Option<String>) -> SemType {
        let Some(raw) = ret else { return SemType::LateBound };
        let name = lc(raw.trim());
        if let Some(elem) = name.strip_suffix("()") {
            return SemType::Array(Box::new(self.lib_return_sem(lib, &Some(elem.to_string()))));
        }
        let value = match name.as_str() {
            "" | "void" | "variant" | "object" | "any" | "idispatch" | "iunknown" => {
                return SemType::LateBound
            }
            "integer" => Some(VbaType::Integer),
            "long" => Some(VbaType::Long),
            "single" => Some(VbaType::Single),
            "double" => Some(VbaType::Double),
            "currency" => Some(VbaType::Currency),
            "date" => Some(VbaType::Date),
            "string" => Some(VbaType::String),
            "boolean" => Some(VbaType::Boolean),
            "byte" => Some(VbaType::Byte),
            "decimal" => Some(VbaType::Decimal),
            _ => None,
        };
        if let Some(v) = value {
            return SemType::Value(v);
        }
        match self.lib_type_ref(lib, &name) {
            Some((l, t)) => match self.libraries[l].types[t].kind {
                LibraryTypeKind::Enum => SemType::Value(VbaType::Long),
                _ => SemType::Lib { lib: l, ty: t },
            },
            None => SemType::Unknown(name),
        }
    }

    fn lib_member(&self, lib: usize, ty: usize, name: &str, access: Access) -> MemberLookup {
        let mut named_somewhere = false;
        for (l, t) in self.lib_chain(lib, ty) {
            let ty_data = &self.libraries[l].types[t];
            let mut named: Vec<usize> = Vec::new();
            for (mi, m) in ty_data.members.iter().enumerate() {
                if m.name == name && m.kind != LibraryMemberKind::Event {
                    named.push(mi);
                    named_somewhere = true;
                }
            }
            if let Some(&mi) = named.iter().find(|&&mi| lib_serves(ty_data.members[mi].kind, access)) {
                let m = &ty_data.members[mi];
                return MemberLookup::Found(MemberHit {
                    resolution: NameResolution::LibraryMember { lib: l, ty: t, member: mi },
                    sem: self.lib_return_sem(l, &m.return_type),
                    callable: lib_callable(m),
                });
            }
        }
        if named_somewhere {
            MemberLookup::MissingAccessor(access)
        } else if self
            .lib_chain(lib, ty)
            .iter()
            .any(|&(l, t)| self.libraries[l].types[t].extensible)
        {
            MemberLookup::LateBound
        } else {
            MemberLookup::NoSuchMember
        }
    }

    // ── Module member lookup ────────────────────────────────────────────────

    /// Public procedures and variables of `module` named `name`, for a use
    /// needing `access`. Procedure results are typed in `module`'s context.
    fn module_member(&self, module: usize, name: &str, access: Access) -> MemberLookup {
        let m = &self.modules[module];
        if let Some(vi) = m.vars.iter().position(|v| v.is_public && v.name == name) {
            return MemberLookup::Found(MemberHit {
                resolution: NameResolution::External { module, decl: ExternalDecl::Var(vi) },
                sem: self.declared_sem(module, &m.vars[vi].ty),
                callable: None,
            });
        }
        let named: Vec<usize> = m
            .procs
            .iter()
            .enumerate()
            .filter(|(_, p)| p.is_public && p.name == name)
            .map(|(i, _)| i)
            .collect();
        if named.is_empty() {
            return MemberLookup::NoSuchMember;
        }
        match named.iter().copied().find(|&i| serves(m.procs[i].kind, access)) {
            Some(pi) => {
                let p = &m.procs[pi];
                // The value of a use is the property's `Get` result (or the
                // function's); an assignment target has no value.
                let sem = match access {
                    Access::Get => self.declared_sem(module, &p.ret),
                    _ => SemType::LateBound,
                };
                MemberLookup::Found(MemberHit {
                    resolution: NameResolution::External { module, decl: ExternalDecl::Proc(pi) },
                    sem,
                    callable: Some(Arity::Known(p.params)),
                })
            }
            None => MemberLookup::MissingAccessor(access),
        }
    }

    fn enum_namespace(&self, module: usize, name: &str) -> Option<MemberHit> {
        let m = &self.modules[module];
        let idx = m.enums.iter().position(|e| e.is_public && e.name == name)?;
        Some(MemberHit {
            resolution: NameResolution::External { module, decl: ExternalDecl::Enum(idx) },
            sem: SemType::Namespace(Qualifier::Enum { module, idx }),
            callable: None,
        })
    }

    /// The library type describing the members a container adds to a control
    /// of `module`, when the module is a UserControl: the extender type of the
    /// library that declares `UserControl`.
    fn extender_type(&self, module: usize) -> Option<(usize, usize)> {
        let base = self.modules[module].base_type.as_ref()?;
        if base.name != "usercontrol" {
            return None;
        }
        let lib = *self.library_by_name.get(base.qualifier.as_deref()?)?;
        let ty = *self.libraries[lib].type_by_name.get("vbcontrolextender")?;
        Some((lib, ty))
    }

    /// A member of a public enum of `module`, named through the module
    /// (`modUtil.Red`, `clsTrip.Done`).
    fn enum_member(&self, module: usize, name: &str) -> Option<MemberHit> {
        let m = &self.modules[module];
        m.enums.iter().enumerate().filter(|(_, e)| e.is_public).find_map(|(ei, e)| {
            let mi = e.members.iter().position(|n| n == name)?;
            Some(MemberHit {
                resolution: NameResolution::External {
                    module,
                    decl: ExternalDecl::EnumMember { enum_idx: ei, member_idx: mi },
                },
                sem: SemType::Value(VbaType::Long),
                callable: None,
            })
        })
    }

    /// A public enum, or a member of one, of `module`.
    fn enum_namespace_or_member(&self, module: usize, name: &str) -> Option<MemberHit> {
        self.enum_namespace(module, name).or_else(|| self.enum_member(module, name))
    }

    fn module_name_hit(&self, module: usize) -> MemberHit {
        let m = &self.modules[module];
        // A form (always) and a class with `VB_PredeclaredId = True` are also a
        // default instance of themselves; a standard module and a plain class
        // name are only a qualifier.
        let sem = if !m.standard && m.predeclared {
            SemType::Class { module }
        } else {
            SemType::Namespace(Qualifier::Module(module))
        };
        MemberHit { resolution: NameResolution::Module(module), sem, callable: None }
    }

    /// Member `name` of the global names library `lib` provides, as the
    /// library name qualifies them (`VBA.IsError`, `VBA.vbCrLf`).
    fn library_member_global(&self, lib: usize, name: &str) -> Option<MemberHit> {
        let l = &self.libraries[lib];
        match l.globals.get(name)? {
            LibGlobal::Object(ty) => Some(MemberHit {
                resolution: NameResolution::LibraryType { lib, ty: *ty },
                sem: SemType::Lib { lib, ty: *ty },
                callable: None,
            }),
            LibGlobal::Member { ty, member } => {
                let m = &l.types[*ty].members[*member];
                Some(MemberHit {
                    resolution: NameResolution::LibraryMember { lib, ty: *ty, member: *member },
                    sem: match l.types[*ty].kind {
                        LibraryTypeKind::Enum => SemType::Value(VbaType::Long),
                        _ => self.lib_return_sem(lib, &m.return_type),
                    },
                    callable: lib_callable(m),
                })
            }
        }
    }

    fn library_global(&self, name: &str) -> Option<MemberHit> {
        // An enum's name qualifies its members.
        for (l, lib) in self.libraries.iter().enumerate() {
            if let Some(&ty) = lib.type_by_name.get(name) {
                if lib.types[ty].kind == LibraryTypeKind::Enum && !lib.globals.contains_key(name) {
                    return Some(MemberHit {
                        resolution: NameResolution::LibraryType { lib: l, ty },
                        sem: SemType::Namespace(Qualifier::LibraryEnum { lib: l, ty }),
                        callable: None,
                    });
                }
            }
        }
        for (l, lib) in self.libraries.iter().enumerate() {
            match lib.globals.get(name) {
                Some(LibGlobal::Object(ty)) => {
                    return Some(MemberHit {
                        resolution: NameResolution::LibraryType { lib: l, ty: *ty },
                        sem: SemType::Lib { lib: l, ty: *ty },
                        callable: None,
                    })
                }
                Some(LibGlobal::Member { ty, member }) => {
                    let m = &lib.types[*ty].members[*member];
                    return Some(MemberHit {
                        resolution: NameResolution::LibraryMember { lib: l, ty: *ty, member: *member },
                        sem: match lib.types[*ty].kind {
                            LibraryTypeKind::Enum => SemType::Value(VbaType::Long),
                            _ => self.lib_return_sem(l, &m.return_type),
                        },
                        callable: lib_callable(m),
                    });
                }
                None => {}
            }
        }
        if let Some(&l) = self.library_by_name.get(name) {
            return Some(MemberHit {
                resolution: NameResolution::Library(l),
                sem: SemType::Namespace(Qualifier::Library(l)),
                callable: None,
            });
        }
        None
    }
}

fn summarize(m: &ModuleData) -> ModuleSummary {
    let name_of = |sym: u32| m.ctx.symbol(sym as usize).name.to_ascii_lowercase();
    let b = &m.raw_bound;
    let declared = |vba: &VbaType, tn: Option<&Option<TypeName>>| Declared {
        vba: vba.clone(),
        name: tn.cloned().flatten(),
    };
    ModuleSummary {
        name: m.module_name.to_ascii_lowercase(),
        display: m.module_name.clone(),
        standard: m.standard,
        predeclared: m.predeclared,
        global_namespace: b.global_namespace,
        default_member: b.default_member.clone(),
        base_type: m.base_type.clone(),
        procs: b
            .procs
            .iter()
            .enumerate()
            .map(|(i, p)| ProcSummary {
                name: name_of(p.sym_id),
                kind: p.kind,
                is_public: p.is_public,
                params: p.params.len(),
                ret: declared(&p.ret_type, b.decl_types.proc_rets.get(i)),
            })
            .collect(),
        vars: b
            .module_vars
            .iter()
            .enumerate()
            .map(|(i, v)| VarSummary {
                name: name_of(v.sym_id),
                is_public: v.is_public,
                ty: declared(&v.vba_type, b.decl_types.module_vars.get(i)),
            })
            .collect(),
        types: b
            .type_decls
            .iter()
            .enumerate()
            .map(|(ti, t)| TypeSummary {
                name: name_of(t.sym_id),
                members: t
                    .members
                    .iter()
                    .enumerate()
                    .map(|(mi, mem)| {
                        let tn = b.decl_types.type_members.get(ti).and_then(|v| v.get(mi));
                        (name_of(mem.sym_id), declared(&mem.vba_type, tn))
                    })
                    .collect(),
            })
            .collect(),
        types_public: b.type_decls.iter().map(|t| t.is_public).collect(),
        enums: b
            .enum_decls
            .iter()
            .map(|e| EnumSummary {
                name: name_of(e.sym_id),
                is_public: e.is_public,
                members: e.members.iter().map(|mem| name_of(mem.sym_id)).collect(),
            })
            .collect(),
        controls: m
            .controls
            .iter()
            .map(|c| ControlSummary {
                name: c.name.to_ascii_lowercase(),
                control_type: c.control_type.clone(),
                is_array: c.is_array,
            })
            .collect(),
    }
}

fn summarize_library(l: &Library) -> LibSummary {
    let types: Vec<LibType> = l.types.iter().map(summarize_lib_type).collect();
    let mut type_by_name = HashMap::new();
    let mut globals = HashMap::new();
    for (ti, t) in l.types.iter().enumerate() {
        type_by_name.entry(lc(&t.name)).or_insert(ti);
    }
    for (ti, t) in l.types.iter().enumerate() {
        if t.is_app_object {
            globals.entry(lc(&t.name)).or_insert(LibGlobal::Object(ti));
        }
    }
    for (ti, t) in l.types.iter().enumerate() {
        let global = matches!(t.kind, LibraryTypeKind::Module | LibraryTypeKind::Enum) || t.is_global;
        if !global {
            continue;
        }
        // The members of a global coclass are those of the interfaces it
        // exposes; each is recorded under the type that declares it.
        for owner in interface_closure(l, &type_by_name, ti) {
            for (mi, m) in l.types[owner].members.iter().enumerate() {
                if m.kind == LibraryMemberKind::Event {
                    continue;
                }
                for key in global_keys(&m.name) {
                    globals.entry(key).or_insert(LibGlobal::Member { ty: owner, member: mi });
                }
            }
        }
    }
    let type_names = l.types.iter().map(|t| t.name.clone()).collect();
    LibSummary { name: lc(&l.name), type_names, types, type_by_name, globals }
}

/// Type `ti` and, transitively, the types of the same library it lists as its
/// default interface, other interfaces or bases, in lookup order.
fn interface_closure(l: &Library, by_name: &HashMap<String, usize>, ti: usize) -> Vec<usize> {
    let mut out = Vec::new();
    let mut stack = vec![ti];
    while let Some(t) = stack.pop() {
        if out.contains(&t) {
            continue;
        }
        out.push(t);
        let ty = &l.types[t];
        let mut next: Vec<usize> = Vec::new();
        for n in ty.default_interface.iter().chain(&ty.interfaces).chain(&ty.extends) {
            if let Some(&i) = by_name.get(&lc(n)) {
                next.push(i);
            }
        }
        stack.extend(next.into_iter().rev());
    }
    out
}

/// The names a library procedure answers to in source. The runtime library
/// spells the `Variant` and `String` forms of a function `_B_var_Name` and
/// `_B_str_Name`; source writes them `Name` and `Name$`.
fn global_keys(member: &str) -> Vec<String> {
    let l = lc(member);
    if let Some(rest) = l.strip_prefix("_b_var_") {
        vec![rest.to_string()]
    } else if let Some(rest) = l.strip_prefix("_b_str_") {
        vec![format!("{rest}$")]
    } else {
        vec![l]
    }
}

fn summarize_lib_type(t: &LibraryType) -> LibType {
    LibType {
        kind: t.kind,
        extensible: t.extensible,
        default_interface: t.default_interface.as_deref().map(lc),
        interfaces: t.interfaces.iter().map(|s| lc(s)).collect(),
        extends: t.extends.iter().map(|s| lc(s)).collect(),
        members: t
            .members
            .iter()
            .map(|m| LibMember {
                name: lc(&m.name),
                kind: m.kind,
                return_type: m.return_type.clone(),
                is_default: m.is_default,
                param_count: m.param_count,
            })
            .collect(),
    }
}

impl ProjectScope for ProjectModel {
    fn sem_of_type_name(&self, from: usize, ty: &TypeName) -> SemType {
        match self.type_ref(from, ty) {
            TypeRef::Udt { module, idx } => SemType::Udt { module, idx },
            TypeRef::Enum { .. } | TypeRef::LibEnum { .. } => SemType::Value(VbaType::Long),
            TypeRef::Class { module } => SemType::Class { module },
            TypeRef::Lib { lib, ty } => SemType::Lib { lib, ty },
            TypeRef::Unknown(n) => SemType::Unknown(n),
            TypeRef::NotLoaded(q) => SemType::LibraryNotLoaded(q),
        }
    }

    fn resolve_type_decl(&self, from: usize, ty: &TypeName) -> NameResolution {
        match self.type_ref(from, ty) {
            TypeRef::Udt { module, idx } => {
                NameResolution::External { module, decl: ExternalDecl::Type(idx) }
            }
            TypeRef::Enum { module, idx } => {
                NameResolution::External { module, decl: ExternalDecl::Enum(idx) }
            }
            TypeRef::Class { module } => NameResolution::Module(module),
            TypeRef::Lib { lib, ty } | TypeRef::LibEnum { lib, ty } => {
                NameResolution::LibraryType { lib, ty }
            }
            TypeRef::Unknown(_) | TypeRef::NotLoaded(_) => NameResolution::Unresolved,
        }
    }

    fn resolve_qualifier(&self, _from: usize, qualifier: &str) -> NameResolution {
        if self.is_project(qualifier) {
            NameResolution::Project
        } else if let Some(&m) = self.module_by_name.get(qualifier) {
            NameResolution::Module(m)
        } else if let Some(&l) = self.library_by_name.get(qualifier) {
            NameResolution::Library(l)
        } else {
            NameResolution::Unresolved
        }
    }

    fn type_name(&self, ty: &SemType) -> String {
        match ty {
            SemType::Class { module } => self.modules[*module].display.clone(),
            SemType::Udt { module, idx } => self.modules[*module].types[*idx].name.clone(),
            SemType::Lib { lib, ty } => format!(
                "{}.{}",
                self.libraries[*lib].name,
                self.libraries[*lib].type_names[*ty]
            ),
            SemType::Value(v) => format!("{v:?}"),
            SemType::Array(inner) => format!("{}()", self.type_name(inner)),
            SemType::Unknown(n) | SemType::LibraryNotLoaded(n) => n.clone(),
            SemType::LateBound => "Variant".to_string(),
            SemType::Namespace(_) | SemType::Unresolved => String::new(),
        }
    }

    fn self_type(&self, from: usize) -> SemType {
        if self.modules[from].standard {
            SemType::LateBound
        } else {
            SemType::Class { module: from }
        }
    }

    fn lookup_module_scope(&self, from: usize, name: &str, access: Access) -> Option<MemberHit> {
        let m = &self.modules[from];
        let Some(idx) = m.controls.iter().position(|c| c.name == name) else {
            // Not a control: the designer's own intrinsic object, or a member
            // of the library type a designer extends.
            let base = m.base_type.as_ref()?;
            // Inside a UserControl, PropertyPage or UserDocument module, the
            // type's name stands for the module's own object (`UserControl.Ambient`).
            if matches!(base.name.as_str(), "usercontrol" | "propertypage" | "userdocument")
                && name == base.name
            {
                if let TypeRef::Lib { lib, ty } = self.type_ref(from, base) {
                    return Some(MemberHit {
                        resolution: NameResolution::LibraryType { lib, ty },
                        sem: SemType::Lib { lib, ty },
                        callable: None,
                    });
                }
            }
            return match self.sem_of_type_name(from, base) {
                SemType::Lib { lib, ty } => match self.lib_member(lib, ty, name, access) {
                    MemberLookup::Found(hit) => Some(hit),
                    _ => None,
                },
                _ => None,
            };
        };
        let c = &m.controls[idx];
        let (q, n) = match c.control_type.split_once('.') {
            Some((q, n)) => (Some(lc(q)), lc(n)),
            None => (None, lc(&c.control_type)),
        };
        let elem = self.sem_of_type_name(from, &TypeName { qualifier: q, name: n });
        let sem = if c.is_array { SemType::Array(Box::new(elem)) } else { elem };
        Some(MemberHit {
            resolution: NameResolution::External { module: from, decl: ExternalDecl::Control(idx) },
            sem,
            callable: None,
        })
    }

    fn lookup_global(&self, _from: usize, name: &str) -> Option<MemberHit> {
        if let Some(r) = self.index.lookup(name) {
            match r.decl {
                ExternalDecl::Proc(pi) => {
                    let p = &self.modules[r.module].procs[pi];
                    return Some(MemberHit {
                        resolution: NameResolution::External { module: r.module, decl: r.decl },
                        sem: self.declared_sem(r.module, &p.ret),
                        callable: Some(Arity::Known(p.params)),
                    });
                }
                ExternalDecl::Var(vi) => {
                    let v = &self.modules[r.module].vars[vi];
                    return Some(MemberHit {
                        resolution: NameResolution::External { module: r.module, decl: r.decl },
                        sem: self.declared_sem(r.module, &v.ty),
                        callable: None,
                    });
                }
                ExternalDecl::Enum(idx) => {
                    return Some(MemberHit {
                        resolution: NameResolution::External { module: r.module, decl: r.decl },
                        sem: SemType::Namespace(Qualifier::Enum { module: r.module, idx }),
                        callable: None,
                    });
                }
                ExternalDecl::EnumMember { .. } => {
                    return Some(MemberHit {
                        resolution: NameResolution::External { module: r.module, decl: r.decl },
                        sem: SemType::Value(VbaType::Long),
                        callable: None,
                    });
                }
                // A type name is not a value, and is never a bare name's
                // referent in an expression.
                ExternalDecl::Type(_) | ExternalDecl::TypeMember { .. } | ExternalDecl::Control(_) => {}
            }
        }
        if let Some(&m) = self.module_by_name.get(name) {
            return Some(self.module_name_hit(m));
        }
        if self.is_project(name) {
            return Some(MemberHit {
                resolution: NameResolution::Project,
                sem: SemType::Namespace(Qualifier::Project),
                callable: None,
            });
        }
        self.library_global(name)
    }

    fn lookup_member(&self, from: usize, base: &SemType, name: &str, access: Access) -> MemberLookup {
        match base {
            SemType::Udt { module, idx } => {
                let t = &self.modules[*module].types[*idx];
                match t.members.iter().position(|(n, _)| n == name) {
                    Some(mi) => MemberLookup::Found(MemberHit {
                        resolution: NameResolution::External {
                            module: *module,
                            decl: ExternalDecl::TypeMember { type_idx: *idx, member_idx: mi },
                        },
                        sem: self.declared_sem(*module, &t.members[mi].1),
                        callable: None,
                    }),
                    None => MemberLookup::NoSuchMember,
                }
            }
            SemType::Class { module } => {
                let m = &self.modules[*module];
                if let Some(ci) = m.controls.iter().position(|c| c.name == name) {
                    if let Some(hit) =
                        self.lookup_module_scope(*module, &m.controls[ci].name, access)
                    {
                        return MemberLookup::Found(hit);
                    }
                }
                match self.module_member(*module, name, access) {
                    MemberLookup::NoSuchMember => {}
                    other => return other,
                }
                if let Some(hit) = self.enum_namespace_or_member(*module, name) {
                    return MemberLookup::Found(hit);
                }
                // A UserControl seen from outside also has the members its
                // container gives every control (position, visibility, ...).
                if from != *module {
                    if let Some((lib, ty)) = self.extender_type(*module) {
                        if let found @ MemberLookup::Found(_) = self.lib_member(lib, ty, name, access) {
                            return found;
                        }
                    }
                }
                match &m.base_type {
                    // A designer's own members were searched; the rest come from
                    // the library type it extends (`VB.Form`).
                    Some(base) => match self.sem_of_type_name(*module, base) {
                        // The module's own class is closed: only a variable
                        // declared with the generic library type (`As Form`)
                        // looks members up at run time.
                        SemType::Lib { lib, ty } => match self.lib_member(lib, ty, name, access) {
                            MemberLookup::LateBound => MemberLookup::NoSuchMember,
                            other => other,
                        },
                        _ => MemberLookup::LibraryNotLoaded(match &base.qualifier {
                            Some(q) => format!("{q}.{}", base.name),
                            None => base.name.clone(),
                        }),
                    },
                    None => MemberLookup::NoSuchMember,
                }
            }
            SemType::Lib { lib, ty } => self.lib_member(*lib, *ty, name, access),
            SemType::Namespace(Qualifier::Module(module)) => {
                let m = &self.modules[*module];
                if m.standard {
                    match self.module_member(*module, name, access) {
                        MemberLookup::NoSuchMember => {}
                        other => return other,
                    }
                }
                match self.enum_namespace_or_member(*module, name) {
                    Some(hit) => MemberLookup::Found(hit),
                    None => MemberLookup::NoSuchMember,
                }
            }
            SemType::Namespace(Qualifier::Project) => {
                if let Some(&m) = self.module_by_name.get(name) {
                    return MemberLookup::Found(self.module_name_hit(m));
                }
                for (i, m) in self.modules.iter().enumerate() {
                    if m.standard || m.global_namespace {
                        match self.module_member(i, name, access) {
                            MemberLookup::NoSuchMember => {}
                            other => return other,
                        }
                    }
                    if let Some(hit) = self.enum_namespace_or_member(i, name) {
                        return MemberLookup::Found(hit);
                    }
                }
                MemberLookup::NoSuchMember
            }
            SemType::Namespace(Qualifier::Library(lib)) => {
                match self.libraries[*lib].type_by_name.get(name) {
                    Some(&ty) => MemberLookup::Found(MemberHit {
                        resolution: NameResolution::LibraryType { lib: *lib, ty },
                        sem: match self.libraries[*lib].types[ty].kind {
                            LibraryTypeKind::Enum => {
                                SemType::Namespace(Qualifier::LibraryEnum { lib: *lib, ty })
                            }
                            _ => SemType::Lib { lib: *lib, ty },
                        },
                        callable: None,
                    }),
                    // A global member of the library, qualified by its name.
                    None => match self.library_member_global(*lib, name) {
                        Some(hit) => MemberLookup::Found(hit),
                        None => MemberLookup::NoSuchMember,
                    },
                }
            }
            SemType::Namespace(Qualifier::LibraryEnum { lib, ty }) => {
                let t = &self.libraries[*lib].types[*ty];
                match t.members.iter().position(|m| m.name == name) {
                    Some(mi) => MemberLookup::Found(MemberHit {
                        resolution: NameResolution::LibraryMember { lib: *lib, ty: *ty, member: mi },
                        sem: SemType::Value(VbaType::Long),
                        callable: None,
                    }),
                    None => MemberLookup::NoSuchMember,
                }
            }
            SemType::Namespace(Qualifier::Enum { module, idx }) => {
                let e = &self.modules[*module].enums[*idx];
                match e.members.iter().position(|n| n == name) {
                    Some(mi) => MemberLookup::Found(MemberHit {
                        resolution: NameResolution::External {
                            module: *module,
                            decl: ExternalDecl::EnumMember { enum_idx: *idx, member_idx: mi },
                        },
                        sem: SemType::Value(VbaType::Long),
                        callable: None,
                    }),
                    None => MemberLookup::NoSuchMember,
                }
            }
            SemType::Value(_)
            | SemType::LateBound
            | SemType::Array(_)
            | SemType::Unknown(_)
            | SemType::LibraryNotLoaded(_)
            | SemType::Unresolved => MemberLookup::NoSuchMember,
        }
    }

    fn default_member(&self, _from: usize, base: &SemType) -> Option<MemberHit> {
        match base {
            SemType::Class { module } => {
                let name = self.modules[*module].default_member.clone()?;
                match self.module_member(*module, &name, Access::Get) {
                    MemberLookup::Found(hit) => Some(hit),
                    _ => None,
                }
            }
            SemType::Lib { lib, ty } => {
                for (l, t) in self.lib_chain(*lib, *ty) {
                    let data = &self.libraries[l].types[t];
                    if let Some(mi) = data.members.iter().position(|m| m.is_default) {
                        let m = &data.members[mi];
                        return Some(MemberHit {
                            resolution: NameResolution::LibraryMember { lib: l, ty: t, member: mi },
                            sem: self.lib_return_sem(l, &m.return_type),
                            callable: lib_callable(m),
                        });
                    }
                }
                None
            }
            _ => None,
        }
    }
}
