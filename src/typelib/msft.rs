//! Describing a type library file (the `MSFT` format) as an engine [`Library`], without the
//! operating system's type-information interfaces.
//!
//! The description matches the one the Windows reader produces from the same library: the same
//! types, members, parameters and spellings.

use super::imports::{guid_key, ImportRequest, ImportResolver, ImportedType};
use super::names::*;
use msft_typelib::{ConstValue, FuncRecord, ResolvedHreftype, TypeDescription, TypeInfoEntry, TypeKind, TypeLib};
use vb6_engine::session::{
    Library, LibraryMember, LibraryMemberKind, LibraryParam, LibraryType, LibraryTypeKind,
};

const FUNCFLAG_FRESTRICTED: u32 = 0x1;
const INVOKE_FUNC: u8 = 1;
const INVOKE_PROPERTYGET: u8 = 2;
const INVOKE_PROPERTYPUT: u8 = 4;
const INVOKE_PROPERTYPUTREF: u8 = 8;
const VARKIND_CONST: i16 = 2;
const DEPTH_LIMIT: usize = 32;
/// The types a constant packed into its own field can have.
const INLINE_VTS: [u16; 14] = [0, 1, 2, 3, 10, 11, 16, 17, 18, 19, 22, 23, 24, 25];

/// The library `data` holds, with types of other libraries named through `imports`.
/// With it, the files of the libraries it imports types from that could not be found: the types
/// taken from them are left undescribed.
pub fn read_library(data: &[u8], imports: &dyn ImportResolver) -> Result<(Library, Vec<String>), String> {
    let lib = TypeLib::parse(data).map_err(|e| e.to_string())?;
    let name = lib.lib_name().unwrap_or_default().to_string();
    let reader = Reader { lib: &lib, library: name.clone(), imports, unresolved: Default::default() };
    let mut types = Vec::new();
    for index in 0..lib.typeinfo_count() {
        let ti = lib.typeinfo(index).map_err(|e| format!("type {index}: {e}"))?;
        if let Some(t) = reader.type_of(&ti)? {
            types.push(t);
        }
    }
    let unresolved = reader.unresolved.into_inner();
    Ok((Library { name, types }, unresolved))
}

/// The GUID of the library `data` holds, lower case without braces.
pub fn library_guid(data: &[u8]) -> Option<String> {
    let lib = TypeLib::parse(data).ok()?;
    lib.lib_guid().map(|g| guid_key(&g.to_string()))
}

/// What a reference type is, as far as parameter passing and spelling need.
struct Target {
    name: String,
    object: bool,
    /// The name is a plain VB6 type (an alias resolved to one), not a type of a library.
    plain: bool,
}

struct Reader<'a, 'd> {
    lib: &'a TypeLib<'d>,
    library: String,
    imports: &'a dyn ImportResolver,
    /// Files of imported libraries that were not found, in the order they were needed.
    unresolved: std::cell::RefCell<Vec<String>>,
}

impl Reader<'_, '_> {
    fn own_name(&self, ti: &TypeInfoEntry<'_>) -> String {
        self.lib.name(ti.name_offset()).unwrap_or_default().to_string()
    }

    /// The type reached through `hreftype`, with aliases followed to what they stand for.
    fn resolve(&self, hreftype: i32, depth: usize) -> Option<Target> {
        match self.lib.resolve_hreftype_full(hreftype)? {
            ResolvedHreftype::Internal(index) => {
                let ti = self.lib.typeinfo(index).ok()?;
                if ti.typekind() == TypeKind::Alias && depth < DEPTH_LIMIT {
                    return match self.lib.type_description(ti.datatype1())? {
                        TypeDescription::UserDefined(next) => self.resolve(next, depth + 1),
                        _ => {
                            let name = self.spell(ti.datatype1(), depth + 1).unwrap_or_else(|| "Variant".into());
                            let vt = (ti.datatype1() as u32 & 0xFFF) as u16;
                            let object = ti.datatype1() < 0 && (vt == VT_DISPATCH || vt == VT_UNKNOWN);
                            Some(Target { name, object, plain: true })
                        }
                    };
                }
                let object = matches!(ti.typekind(), TypeKind::Interface | TypeKind::Dispatch | TypeKind::CoClass);
                Some(Target { name: self.own_name(&ti), object, plain: false })
            }
            ResolvedHreftype::External { imp_info, imp_file } => {
                let file = imp_file?;
                let library_guid = guid_key(&self.lib.guid(file.guid_offset())?.to_string());
                let by_guid = imp_info.flags() & 0x10000 != 0;
                let type_guid = if by_guid { self.lib.guid(imp_info.guid_offset()).map(|g| guid_key(&g.to_string())) } else { None };
                let type_index = if by_guid { None } else { usize::try_from(imp_info.guid_offset()).ok() };
                let file_name = file.filename().unwrap_or_default();
                let found = self.imports.resolve(&ImportRequest {
                    file: file_name,
                    library_guid: &library_guid,
                    type_guid: type_guid.as_deref(),
                    type_index,
                });
                let Some(found) = found else {
                    let mut missing = self.unresolved.borrow_mut();
                    if !missing.iter().any(|m| m == file_name) {
                        missing.push(file_name.to_string());
                    }
                    return None;
                };
                Some(self.qualified(found))
            }
        }
    }

    fn qualified(&self, found: ImportedType) -> Target {
        let plain = found.library.is_empty();
        let name = if found.library.eq_ignore_ascii_case(&self.library) || plain {
            found.name
        } else {
            format!("{}.{}", found.library, found.name)
        };
        Target { name, object: found.object, plain }
    }

    /// The name of the type a reference names, qualified when it belongs to another library.
    fn reference_name(&self, hreftype: i32) -> Option<String> {
        match self.lib.resolve_hreftype_full(hreftype)? {
            ResolvedHreftype::Internal(index) => Some(self.own_name(&self.lib.typeinfo(index).ok()?)),
            ResolvedHreftype::External { .. } => Some(self.resolve(hreftype, 0)?.name),
        }
    }

    /// The VB6 spelling of the type `datatype` describes, with pointers collapsed.
    fn spell(&self, datatype: i32, depth: usize) -> Option<String> {
        if depth > DEPTH_LIMIT {
            return None;
        }
        match self.lib.type_description(datatype)? {
            TypeDescription::Base(vt) => plain_type_name(vt).map(str::to_string),
            TypeDescription::Pointer(inner) => self.spell(inner, depth + 1),
            TypeDescription::SafeArray(inner) => self.spell(inner, depth + 1).map(|e| array_of(&e)),
            TypeDescription::CArray(array) => self.spell(array.element_type(), depth + 1).map(|e| array_of(&e)),
            TypeDescription::UserDefined(href) => Some(self.resolve(href, 0)?.name),
        }
    }

    /// The spelling of a parameter's type and whether it is passed by reference.
    fn param_type(&self, datatype: i32) -> (Option<String>, bool) {
        let Some(TypeDescription::Pointer(inner)) = self.lib.type_description(datatype) else {
            return (self.spell(datatype, 0), false);
        };
        let by_ref = match self.lib.type_description(inner) {
            // A pointer to an interface, class or dispatch type is an object passed by value.
            Some(TypeDescription::UserDefined(href)) => self.resolve(href, 0).is_none_or(|t| !t.object),
            Some(TypeDescription::Base(vt)) if vt == VT_DISPATCH || vt == VT_UNKNOWN => false,
            _ => true,
        };
        (self.spell(datatype, 0), by_ref)
    }

    fn type_of(&self, ti: &TypeInfoEntry<'_>) -> Result<Option<LibraryType>, String> {
        let kind = match ti.typekind() {
            TypeKind::Enum => LibraryTypeKind::Enum,
            TypeKind::Record | TypeKind::Union => LibraryTypeKind::Record,
            TypeKind::Module => LibraryTypeKind::Module,
            TypeKind::Interface | TypeKind::Dispatch => LibraryTypeKind::Interface,
            TypeKind::CoClass => LibraryTypeKind::Coclass,
            // An alias is not a type of its own: uses of it are replaced by what it stands for.
            TypeKind::Alias | TypeKind::Unknown(_) => return Ok(None),
        };
        let flags = ti.flags();
        let mut ty = LibraryType {
            name: self.own_name(ti),
            kind,
            is_app_object: flags & TYPEFLAG_FAPPOBJECT as u32 != 0,
            is_global: flags & TYPEFLAG_FAPPOBJECT as u32 != 0,
            is_control: flags & TYPEFLAG_FCONTROL as u32 != 0,
            ..LibraryType::default()
        };
        if kind == LibraryTypeKind::Coclass {
            self.coclass_parts(ti, &mut ty)?;
        } else if ti.typekind() == TypeKind::Dispatch {
            // A dispatch interface always derives from `IDispatch`: the library's own, when it
            // declares one, else the OLE automation library's.
            ty.extends.push(self.dispatch_base());
        } else if kind == LibraryTypeKind::Interface && ti.cimpltypes() > 0 {
            // An interface derives from at most one other: its reference is stored directly.
            if let Some(base) = self.reference_name(ti.datatype1()) {
                ty.extends.push(base);
            }
        }
        self.members_of(ti, kind, &mut ty.members);
        Ok(Some(ty))
    }

    fn dispatch_base(&self) -> String {
        const IDISPATCH: &str = "00020400-0000-0000-c000-000000000046";
        let declared = (0..self.lib.typeinfo_count()).filter_map(|i| self.lib.typeinfo(i).ok()).any(|t| {
            self.lib.guid(t.guid_offset()).is_some_and(|g| guid_key(&g.to_string()) == IDISPATCH)
        });
        if declared || self.library.eq_ignore_ascii_case("stdole") {
            "IDispatch".to_string()
        } else {
            "stdole.IDispatch".to_string()
        }
    }

    fn coclass_parts(&self, ti: &TypeInfoEntry<'_>, ty: &mut LibraryType) -> Result<(), String> {
        let mut default_interface: Option<String> = None;
        let mut interfaces: Vec<String> = Vec::new();
        let mut default_source: Option<usize> = None;
        let mut sources: Vec<usize> = Vec::new();
        for record in self.lib.impl_types(ti) {
            let href = record.reftype();
            if record.is_source() {
                if let Some(ResolvedHreftype::Internal(index)) = self.lib.resolve_hreftype_full(href) {
                    if record.is_default() {
                        default_source = Some(index);
                    } else {
                        sources.push(index);
                    }
                }
                continue;
            }
            let Some(name) = self.reference_name(href) else { continue };
            if record.is_default() && default_interface.is_none() {
                default_interface = Some(name);
            } else {
                interfaces.push(name);
            }
        }
        ty.default_interface = default_interface;
        ty.interfaces = interfaces;
        // The events a class raises are the members of its source interfaces.
        for index in default_source.into_iter().chain(sources) {
            let source = self.lib.typeinfo(index).map_err(|e| e.to_string())?;
            let mut events = Vec::new();
            self.members_of(&source, LibraryTypeKind::Interface, &mut events);
            for mut event in events {
                event.kind = LibraryMemberKind::Event;
                event.is_default = false;
                ty.members.push(event);
            }
        }
        Ok(())
    }

    fn members_of(&self, ti: &TypeInfoEntry<'_>, kind: LibraryTypeKind, out: &mut Vec<LibraryMember>) {
        self.members_of_at(ti, kind, out, 0, ti.typekind());
    }

    /// The members of a type. A dispatch interface lists the members it inherits from the
    /// interfaces of the library it derives from before its own.
    fn members_of_at(
        &self,
        ti: &TypeInfoEntry<'_>,
        kind: LibraryTypeKind,
        out: &mut Vec<LibraryMember>,
        depth: usize,
        owner: TypeKind,
    ) {
        if ti.typekind() == TypeKind::Dispatch && ti.cimpltypes() > 0 && depth < DEPTH_LIMIT {
            if let Some(ResolvedHreftype::Internal(base)) = self.lib.resolve_hreftype_full(ti.datatype1()) {
                if let Ok(base) = self.lib.typeinfo(base) {
                    if matches!(base.typekind(), TypeKind::Dispatch | TypeKind::Interface) {
                        self.members_of_at(&base, kind, out, depth + 1, owner);
                    }
                }
            }
        }
        for (index, func) in self.lib.funcs(ti).enumerate() {
            if let Some(member) = self.function(ti, index, &func, owner) {
                out.push(member);
            }
        }
        for (index, var) in self.lib.vars(ti).enumerate() {
            out.push(self.variable(ti, index, &var, kind));
        }
    }

    fn function(&self, ti: &TypeInfoEntry<'_>, index: usize, f: &FuncRecord<'_>, owner: TypeKind) -> Option<LibraryMember> {
        if f.flags() & FUNCFLAG_FRESTRICTED != 0 {
            return None;
        }
        let memid = self.lib.func_memid(ti, index).unwrap_or(0);
        let name = self.lib.func_name(ti, index).unwrap_or_default().to_string();
        let names = self.parameter_names(ti, index, f, owner);
        let all: Vec<_> = self.lib.params(f).collect();
        let last_visible = all
            .iter()
            .rposition(|p| (p.flags() & 0xFFFF) as u16 & (PARAMFLAG_FLCID | PARAMFLAG_FRETVAL) == 0);
        let mut params: Vec<LibraryParam> = Vec::new();
        let mut result: Option<String> = None;
        let mut returned_through_param = false;
        for (i, p) in all.iter().enumerate() {
            let flags = (p.flags() & 0xFFFF) as u16;
            if flags & PARAMFLAG_FLCID != 0 {
                continue;
            }
            if flags & PARAMFLAG_FRETVAL != 0 {
                result = self.spell(p.datatype(), 0);
                returned_through_param = true;
                continue;
            }
            let (type_name, by_ref) = self.param_type(p.datatype());
            let optional = flags & PARAMFLAG_FOPT != 0;
            let default = if flags & PARAMFLAG_FHASDEFAULT != 0 {
                f.default_value_offset(i).and_then(|o| self.default_constant(o)).and_then(|v| default_text(&v))
            } else {
                None
            };
            let param_array = f.nroargs() == -1 && Some(i) == last_visible;
            params.push(LibraryParam {
                name: names.get(i).cloned().unwrap_or_default(),
                type_name,
                by_ref,
                optional: optional || param_array,
                default,
                param_array,
            });
        }
        if !returned_through_param {
            let valueless = self
                .lib
                .type_description(f.datatype())
                .is_some_and(|d| matches!(d, TypeDescription::Base(vt) if is_valueless(vt)));
            if !valueless {
                result = self.spell(f.datatype(), 0);
            }
        }
        let kind = match f.invoke_kind() {
            INVOKE_FUNC => LibraryMemberKind::Method,
            INVOKE_PROPERTYGET => LibraryMemberKind::PropertyGet,
            INVOKE_PROPERTYPUT => LibraryMemberKind::PropertyLet,
            INVOKE_PROPERTYPUTREF => LibraryMemberKind::PropertySet,
            _ => LibraryMemberKind::Method,
        };
        let flags = f.flags() & 0xFFFF;
        Some(LibraryMember {
            name,
            kind,
            return_type: result,
            is_default: memid == 0 && matches!(kind, LibraryMemberKind::PropertyGet | LibraryMemberKind::Method),
            param_count: Some(params.len()),
            params,
            dispid: Some(memid),
            member_flags: if flags != 0 { Some(flags) } else { None },
            value: None,
        })
    }

    /// The names of a function's parameters, by position. A parameter of a property assignment
    /// the file does not name takes the name of the getter's parameter in the same position:
    /// seen as a vtable interface the getter's parameters are all of them (its returned value
    /// included), seen as a dispatch interface only those a caller passes.
    fn parameter_names(&self, ti: &TypeInfoEntry<'_>, index: usize, f: &FuncRecord<'_>, owner: TypeKind) -> Vec<String> {
        let mut names: Vec<String> = self
            .lib
            .params(f)
            .map(|p| self.lib.name(p.name_offset()).unwrap_or_default().to_string())
            .collect();
        let assigns = matches!(f.invoke_kind(), INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF);
        if !assigns || names.iter().all(|n| !n.is_empty()) {
            return names;
        }
        let memid = self.lib.func_memid(ti, index);
        let getter_names: Option<Vec<String>> = self.lib.funcs(ti).enumerate().find_map(|(i, getter)| {
            if getter.invoke_kind() != INVOKE_PROPERTYGET || self.lib.func_memid(ti, i) != memid {
                return None;
            }
            Some(
                self.lib
                    .params(&getter)
                    .filter(|p| {
                        owner == TypeKind::Interface
                            || (p.flags() & 0xFFFF) as u16 & (PARAMFLAG_FLCID | PARAMFLAG_FRETVAL) == 0
                    })
                    .map(|p| self.lib.name(p.name_offset()).unwrap_or_default().to_string())
                    .collect(),
            )
        });
        if let Some(getter_names) = getter_names {
            for (j, name) in names.iter_mut().enumerate() {
                if name.is_empty() {
                    if let Some(from_getter) = getter_names.get(j) {
                        *name = from_getter.clone();
                    }
                }
            }
        }
        names
    }

    /// The value of a parameter default, when it is one the system would read. A value packed
    /// into its own field keeps the bits of its type in the low 26 bits.
    fn default_constant<'v>(&'v self, offs_value: i32) -> Option<ConstValue<'v>> {
        if offs_value < 0 {
            let vt = (((offs_value as u32) & 0x7C00_0000) >> 26) as u16;
            let bits = (offs_value as u32) & 0x03FF_FFFF;
            match vt {
                4 => return Some(ConstValue::R4(f32::from_bits(bits))),
                5 => return Some(ConstValue::R8(f64::from_bits(bits as u64))),
                _ if !INLINE_VTS.contains(&vt) => return None,
                _ => {}
            }
        }
        self.lib.const_value(offs_value)
    }

    fn variable(
        &self,
        ti: &TypeInfoEntry<'_>,
        index: usize,
        v: &msft_typelib::VarRecord<'_>,
        kind: LibraryTypeKind,
    ) -> LibraryMember {
        let memid = self.lib.var_memid(ti, index).unwrap_or(0);
        let name = self.lib.var_name(ti, index).unwrap_or_default().to_string();
        let is_const = v.varkind() == VARKIND_CONST;
        let value = if is_const { self.lib.const_value(v.offs_value()).and_then(|c| integer_of(&c)) } else { None };
        let constant = is_const || kind == LibraryTypeKind::Enum;
        let type_name = if constant && kind == LibraryTypeKind::Enum { None } else { self.spell(v.datatype(), 0) };
        let flags = v.flags() & 0xFFFF;
        LibraryMember {
            name,
            kind: if constant { LibraryMemberKind::Constant } else { LibraryMemberKind::Field },
            return_type: type_name,
            is_default: memid == 0 && !constant && kind != LibraryTypeKind::Record,
            param_count: None,
            params: Vec::new(),
            dispid: if constant || kind == LibraryTypeKind::Record { None } else { Some(memid) },
            member_flags: if flags != 0 { Some(flags) } else { None },
            value,
        }
    }
}

/// The integer a constant's value holds, when it holds one.
fn integer_of(value: &ConstValue<'_>) -> Option<i64> {
    Some(match value {
        ConstValue::I2(n) => *n as i64,
        ConstValue::I4(n) | ConstValue::Int(n) | ConstValue::Error(n) => *n as i64,
        ConstValue::UI1(n) => *n as i64,
        ConstValue::UI2(n) => *n as i64,
        ConstValue::UI4(n) | ConstValue::UInt(n) => *n as i64,
        ConstValue::I8(n) => *n,
        ConstValue::Bool(b) => {
            if *b {
                -1
            } else {
                0
            }
        }
        _ => return None,
    })
}

/// A parameter's default value as VB source text.
fn default_text(value: &ConstValue<'_>) -> Option<String> {
    Some(match value {
        ConstValue::Bstr(bytes) => {
            let (text, _, _) = encoding_rs::WINDOWS_1252.decode(bytes);
            format!("\"{}\"", text.replace('"', "\"\""))
        }
        ConstValue::NullBstr => "\"\"".to_string(),
        ConstValue::Bool(b) => (if *b { "True" } else { "False" }).to_string(),
        ConstValue::R4(n) => n.to_string(),
        ConstValue::R8(n) => n.to_string(),
        ConstValue::Empty | ConstValue::Null => return None,
        other => integer_of(other)?.to_string(),
    })
}

/// The type `index` of `data`'s library as another library sees it: an alias stands for what it
/// names.
pub fn imported_type(lib: &TypeLib<'_>, index: usize) -> Option<ImportedType> {
    struct NoImports;
    impl ImportResolver for NoImports {
        fn resolve(&self, _: &ImportRequest<'_>) -> Option<ImportedType> {
            None
        }
    }
    let library = lib.lib_name()?.to_string();
    let reader = Reader { lib, library: library.clone(), imports: &NoImports, unresolved: Default::default() };
    let target = reader.resolve((index * TypeInfoEntry::SIZE) as i32, 0)?;
    Some(ImportedType { library: if target.plain { String::new() } else { library }, name: target.name, object: target.object })
}

/// Every type of `data`'s library as an importing library sees it, with its index and, when it
/// has one, its GUID (lower case, without braces).
pub fn exported_types(data: &[u8]) -> Result<Vec<(usize, Option<String>, ImportedType)>, String> {
    let lib = TypeLib::parse(data).map_err(|e| e.to_string())?;
    let mut out = Vec::new();
    for index in 0..lib.typeinfo_count() {
        let ti = lib.typeinfo(index).map_err(|e| e.to_string())?;
        let guid = lib.guid(ti.guid_offset()).map(|g| guid_key(&g.to_string()));
        if let Some(found) = imported_type(&lib, index) {
            out.push((index, guid, found));
        }
    }
    Ok(out)
}
