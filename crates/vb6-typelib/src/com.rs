//! Reading a type library through the automation type-information interfaces
//! and describing it as an engine [`Library`].

use super::locate::{file_of_registered, registered_path};
use super::names::*;
use super::registry::registered_value;
use super::{TypeLibraryError, TypeLibraryRequest};
use vb6_engine::session::{
    Library, LibraryMember, LibraryMemberKind, LibraryParam, LibraryType, LibraryTypeKind,
};
use windows::core::{BSTR, GUID, HSTRING};
use windows::Win32::System::Com::{
    CoInitializeEx, CoUninitialize, ELEMDESC, FUNCDESC, ITypeInfo, ITypeLib, TYPEDESC,
    COINIT_MULTITHREADED, INVOKE_FUNC, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT,
    INVOKE_PROPERTYPUTREF, TKIND_ALIAS, TKIND_COCLASS, TKIND_DISPATCH, TKIND_ENUM,
    TKIND_INTERFACE, TKIND_MODULE, TKIND_RECORD, TKIND_UNION, VAR_CONST,
};
use windows::Win32::System::Ole::{LoadTypeLibEx, REGKIND_NONE};

const FUNCFLAG_FRESTRICTED: u16 = 0x1;

/// Initialises COM for the calling thread for as long as it lives.
struct Apartment(bool);

impl Apartment {
    fn enter() -> Apartment {
        // A thread already in another apartment model is fine: the call fails and nothing is
        // undone on drop.
        Apartment(unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) }.is_ok())
    }
}

impl Drop for Apartment {
    fn drop(&mut self) {
        if self.0 {
            unsafe { CoUninitialize() };
        }
    }
}

/// Read the library `request` names.
pub fn read(request: &TypeLibraryRequest) -> Result<Library, TypeLibraryError> {
    let _apartment = Apartment::enter();
    let lib = open(request)?;
    unsafe { describe(&lib) }.map_err(|e| TypeLibraryError::Unreadable {
        what: request.display_name(),
        reason: e.to_string(),
    })
}

fn guid_of(text: &str) -> Option<GUID> {
    let text = text.trim().trim_matches(|c| c == '{' || c == '}');
    GUID::try_from(text).ok()
}

/// Open the library from the registry, else from the files the request names.
fn open(request: &TypeLibraryRequest) -> Result<ITypeLib, TypeLibraryError> {
    let wanted = guid_of(&request.guid);
    let mut tried: Vec<String> = Vec::new();
    let mut last_reason = String::new();

    let mut candidates: Vec<String> = Vec::new();
    if let Some(value) = registered_value(&request.guid, request.major, request.minor, request.lcid) {
        // The registered value may end in a resource index; load it whole, and fall back to the file.
        candidates.push(registered_path(&value).to_string_lossy().into_owned());
        candidates.push(file_of_registered(&value).to_string_lossy().into_owned());
    }
    for path in &request.paths {
        candidates.push(path.to_string_lossy().into_owned());
    }
    candidates.dedup();

    for candidate in candidates {
        if tried.contains(&candidate) {
            continue;
        }
        tried.push(candidate.clone());
        if !exists(&candidate) {
            continue;
        }
        match unsafe { LoadTypeLibEx(&HSTRING::from(candidate.as_str()), REGKIND_NONE) } {
            Ok(lib) => {
                if let (Some(wanted), Ok(attr)) = (wanted, unsafe { lib.GetLibAttr() }) {
                    let found = unsafe { (*attr).guid };
                    unsafe { lib.ReleaseTLibAttr(attr) };
                    if found != wanted {
                        last_reason = format!("{candidate} holds a different library");
                        continue;
                    }
                }
                return Ok(lib);
            }
            Err(e) => last_reason = format!("{candidate}: {e}"),
        }
    }
    Err(TypeLibraryError::NotFound { what: request.display_name(), tried, reason: last_reason })
}

/// Whether the file a path names exists; the path may end in a resource index (`file.dll`).
fn exists(candidate: &str) -> bool {
    file_of_registered(candidate).exists()
}

fn bstr(f: impl FnOnce(*mut BSTR) -> windows::core::Result<()>) -> Option<String> {
    let mut out = BSTR::new();
    f(&mut out).ok()?;
    Some(out.to_string())
}

unsafe fn describe(lib: &ITypeLib) -> windows::core::Result<Library> {
    let name = bstr(|p| lib.GetDocumentation(-1, Some(p), None, std::ptr::null_mut(), None)).unwrap_or_default();
    let reader = Reader { library: name.clone() };
    let mut types = Vec::new();
    for index in 0..lib.GetTypeInfoCount() {
        let info = lib.GetTypeInfo(index)?;
        if let Some(t) = reader.type_of(&info)? {
            types.push(t);
        }
    }
    Ok(Library { name, types })
}

/// What a reference type is, as far as parameter passing and spelling need.
struct Target {
    name: String,
    object: bool,
}

struct Reader {
    /// The name of the library being read; types of other libraries are qualified with theirs.
    library: String,
}

impl Reader {
    unsafe fn type_name(&self, info: &ITypeInfo) -> windows::core::Result<String> {
        let name = self.own_name(info);
        let mut owner: Option<ITypeLib> = None;
        let mut index = 0u32;
        if info.GetContainingTypeLib(&mut owner, &mut index).is_ok() {
            if let Some(owner) = owner {
                let owner_name =
                    bstr(|p| owner.GetDocumentation(-1, Some(p), None, std::ptr::null_mut(), None)).unwrap_or_default();
                if !owner_name.is_empty() && !owner_name.eq_ignore_ascii_case(&self.library) {
                    return Ok(format!("{owner_name}.{name}"));
                }
            }
        }
        Ok(name)
    }

    /// The reference type `href` of `info`, with aliases followed to what they stand for.
    unsafe fn resolve(&self, info: &ITypeInfo, href: u32) -> windows::core::Result<Target> {
        let mut target = info.GetRefTypeInfo(href)?;
        let mut guard = 0;
        loop {
            let attr = target.GetTypeAttr()?;
            let kind = (*attr).typekind;
            let alias = (*attr).tdescAlias;
            target.ReleaseTypeAttr(attr);
            if kind == TKIND_ALIAS && guard < 32 {
                guard += 1;
                if alias.vt.0 == VT_USERDEFINED {
                    target = target.GetRefTypeInfo(alias.Anonymous.hreftype)?;
                    continue;
                }
                // An alias of a plain type: spell what it stands for.
                let spelled = self.spell(&target, &alias)?.unwrap_or_else(|| "Variant".to_string());
                let object = alias.vt.0 == VT_DISPATCH || alias.vt.0 == VT_UNKNOWN;
                return Ok(Target { name: spelled, object });
            }
            let object = kind == TKIND_INTERFACE || kind == TKIND_DISPATCH || kind == TKIND_COCLASS;
            return Ok(Target { name: self.type_name(&target)?, object });
        }
    }

    /// The VB6 spelling of `t`, with pointers collapsed.
    unsafe fn spell(&self, info: &ITypeInfo, t: &TYPEDESC) -> windows::core::Result<Option<String>> {
        let vt = t.vt.0;
        Ok(match vt {
            VT_PTR => self.spell(info, &*t.Anonymous.lptdesc)?,
            VT_SAFEARRAY => self.spell(info, &*t.Anonymous.lptdesc)?.map(|e| array_of(&e)),
            VT_CARRAY => self.spell(info, &(*t.Anonymous.lpadesc).tdescElem)?.map(|e| array_of(&e)),
            VT_USERDEFINED => Some(self.resolve(info, t.Anonymous.hreftype)?.name),
            _ => plain_type_name(vt).map(str::to_string),
        })
    }

    /// The spelling of a parameter's type and whether it is passed by reference.
    unsafe fn param_type(&self, info: &ITypeInfo, t: &TYPEDESC) -> windows::core::Result<(Option<String>, bool)> {
        if t.vt.0 != VT_PTR {
            return Ok((self.spell(info, t)?, false));
        }
        let inner = &*t.Anonymous.lptdesc;
        let by_ref = match inner.vt.0 {
            // A pointer to an interface, class or dispatch type is an object passed by value.
            VT_USERDEFINED => !self.resolve(info, inner.Anonymous.hreftype)?.object,
            VT_DISPATCH | VT_UNKNOWN => false,
            _ => true,
        };
        Ok((self.spell(info, t)?, by_ref))
    }

    unsafe fn type_of(&self, info: &ITypeInfo) -> windows::core::Result<Option<LibraryType>> {
        let attr = info.GetTypeAttr()?;
        let a = *attr;
        info.ReleaseTypeAttr(attr);
        let kind = match a.typekind {
            k if k == TKIND_ENUM => LibraryTypeKind::Enum,
            k if k == TKIND_RECORD || k == TKIND_UNION => LibraryTypeKind::Record,
            k if k == TKIND_MODULE => LibraryTypeKind::Module,
            k if k == TKIND_INTERFACE || k == TKIND_DISPATCH => LibraryTypeKind::Interface,
            k if k == TKIND_COCLASS => LibraryTypeKind::Coclass,
            // An alias is not a type of its own: uses of it are replaced by what it stands for.
            k if k == TKIND_ALIAS => return Ok(None),
            _ => return Ok(None),
        };
        let name = self.own_name(info);
        let mut ty = LibraryType {
            name,
            kind,
            is_app_object: a.wTypeFlags & TYPEFLAG_FAPPOBJECT != 0,
            is_global: a.wTypeFlags & TYPEFLAG_FAPPOBJECT != 0,
            is_control: a.wTypeFlags & TYPEFLAG_FCONTROL != 0,
            ..LibraryType::default()
        };
        if kind == LibraryTypeKind::Coclass {
            self.coclass_parts(info, a.cImplTypes, &mut ty)?;
        } else if kind == LibraryTypeKind::Interface {
            for i in 0..a.cImplTypes as u32 {
                let href = info.GetRefTypeOfImplType(i)?;
                let base = info.GetRefTypeInfo(href)?;
                ty.extends.push(self.type_name(&base)?);
            }
        }
        self.members_of(info, &a, kind, &mut ty.members)?;
        Ok(Some(ty))
    }

    /// The type's own name, never qualified.
    unsafe fn own_name(&self, info: &ITypeInfo) -> String {
        bstr(|p| info.GetDocumentation(-1, Some(p), None, std::ptr::null_mut(), None)).unwrap_or_default()
    }

    unsafe fn coclass_parts(&self, info: &ITypeInfo, count: u16, ty: &mut LibraryType) -> windows::core::Result<()> {
        let mut default_interface: Option<String> = None;
        let mut interfaces: Vec<String> = Vec::new();
        let mut default_source: Option<ITypeInfo> = None;
        let mut sources: Vec<ITypeInfo> = Vec::new();
        for i in 0..count as u32 {
            let flags = info.GetImplTypeFlags(i)?.0;
            let target = info.GetRefTypeInfo(info.GetRefTypeOfImplType(i)?)?;
            if flags & IMPLTYPEFLAG_FSOURCE != 0 {
                if flags & IMPLTYPEFLAG_FDEFAULT != 0 {
                    default_source = Some(target);
                } else {
                    sources.push(target);
                }
            } else {
                let name = self.type_name(&target)?;
                if flags & IMPLTYPEFLAG_FDEFAULT != 0 && default_interface.is_none() {
                    default_interface = Some(name);
                } else {
                    interfaces.push(name);
                }
            }
        }
        ty.default_interface = default_interface;
        ty.interfaces = interfaces;
        // The events a class raises are the members of its source interfaces.
        for source in default_source.into_iter().chain(sources) {
            let attr = source.GetTypeAttr()?;
            let a = *attr;
            source.ReleaseTypeAttr(attr);
            let mut events = Vec::new();
            self.members_of(&source, &a, LibraryTypeKind::Interface, &mut events)?;
            for mut event in events {
                event.kind = LibraryMemberKind::Event;
                event.is_default = false;
                ty.members.push(event);
            }
        }
        Ok(())
    }

    unsafe fn members_of(
        &self,
        info: &ITypeInfo,
        a: &windows::Win32::System::Com::TYPEATTR,
        kind: LibraryTypeKind,
        out: &mut Vec<LibraryMember>,
    ) -> windows::core::Result<()> {
        for i in 0..a.cFuncs as u32 {
            let desc = info.GetFuncDesc(i)?;
            let member = self.function(info, &*desc);
            info.ReleaseFuncDesc(desc);
            if let Some(member) = member? {
                out.push(member);
            }
        }
        for i in 0..a.cVars as u32 {
            let desc = info.GetVarDesc(i)?;
            let var = *desc;
            let member = self.variable(info, &var, kind);
            info.ReleaseVarDesc(desc);
            if let Some(member) = member? {
                out.push(member);
            }
        }
        Ok(())
    }

    unsafe fn names_of(&self, info: &ITypeInfo, memid: i32, params: usize) -> Vec<String> {
        let mut names = vec![BSTR::new(); params + 1];
        let mut count = 0u32;
        if info.GetNames(memid, &mut names, &mut count).is_err() {
            return vec![String::new(); params + 1];
        }
        let mut out: Vec<String> = names.iter().map(|n| n.to_string()).collect();
        out.truncate(params + 1);
        out
    }

    unsafe fn function(&self, info: &ITypeInfo, f: &FUNCDESC) -> windows::core::Result<Option<LibraryMember>> {
        if f.wFuncFlags.0 & FUNCFLAG_FRESTRICTED != 0 {
            return Ok(None);
        }
        let count = f.cParams.max(0) as usize;
        let names = self.names_of(info, f.memid, count);
        let name = names.first().cloned().unwrap_or_default();
        let elems: &[ELEMDESC] = if count == 0 { &[] } else { std::slice::from_raw_parts(f.lprgelemdescParam, count) };

        let mut params: Vec<LibraryParam> = Vec::new();
        let mut result: Option<String> = None;
        let mut returned_through_param = false;
        let last_visible = elems.iter().rposition(|elem| {
            (elem.Anonymous.paramdesc.wParamFlags.0 as u16) & (PARAMFLAG_FLCID | PARAMFLAG_FRETVAL) == 0
        });
        for (i, elem) in elems.iter().enumerate() {
            let flags = elem.Anonymous.paramdesc.wParamFlags.0 as u16;
            if flags & PARAMFLAG_FLCID != 0 {
                continue;
            }
            if flags & PARAMFLAG_FRETVAL != 0 {
                result = self.spell(info, &elem.tdesc)?;
                returned_through_param = true;
                continue;
            }
            let (type_name, by_ref) = self.param_type(info, &elem.tdesc)?;
            let optional = flags & PARAMFLAG_FOPT != 0;
            let default = if flags & PARAMFLAG_FHASDEFAULT != 0 {
                let ex = elem.Anonymous.paramdesc.pparamdescex;
                if ex.is_null() { None } else { default_text(&(*ex).varDefaultValue) }
            } else {
                None
            };
            let param_array = f.cParamsOpt == -1 && Some(i) == last_visible;
            params.push(LibraryParam {
                name: names.get(i + 1).cloned().unwrap_or_default(),
                type_name,
                by_ref,
                optional: optional || param_array,
                default,
                param_array,
            });
        }
        if !returned_through_param {
            let vt = f.elemdescFunc.tdesc.vt.0;
            if !is_valueless(vt) {
                result = self.spell(info, &f.elemdescFunc.tdesc)?;
            }
        }
        let kind = match f.invkind {
            k if k == INVOKE_FUNC => LibraryMemberKind::Method,
            k if k == INVOKE_PROPERTYGET => LibraryMemberKind::PropertyGet,
            k if k == INVOKE_PROPERTYPUT => LibraryMemberKind::PropertyLet,
            k if k == INVOKE_PROPERTYPUTREF => LibraryMemberKind::PropertySet,
            _ => LibraryMemberKind::Method,
        };
        let flags = f.wFuncFlags.0;
        Ok(Some(LibraryMember {
            name,
            kind,
            return_type: result,
            is_default: f.memid == 0 && matches!(kind, LibraryMemberKind::PropertyGet | LibraryMemberKind::Method),
            param_count: Some(params.len()),
            params,
            dispid: Some(f.memid),
            member_flags: if flags != 0 { Some(flags as u32) } else { None },
            value: None,
        }))
    }

    unsafe fn variable(
        &self,
        info: &ITypeInfo,
        v: &windows::Win32::System::Com::VARDESC,
        kind: LibraryTypeKind,
    ) -> windows::core::Result<Option<LibraryMember>> {
        let names = self.names_of(info, v.memid, 0);
        let name = names.first().cloned().unwrap_or_default();
        let is_const = v.varkind == VAR_CONST;
        let value = if is_const {
            let value = v.Anonymous.lpvarValue;
            if value.is_null() { None } else { integer_of(&*value) }
        } else {
            None
        };
        let constant = is_const || kind == LibraryTypeKind::Enum;
        let type_name = if constant && kind == LibraryTypeKind::Enum {
            None
        } else {
            self.spell(info, &v.elemdescVar.tdesc)?
        };
        Ok(Some(LibraryMember {
            name,
            kind: if constant { LibraryMemberKind::Constant } else { LibraryMemberKind::Field },
            return_type: type_name,
            is_default: v.memid == 0 && !constant && kind != LibraryTypeKind::Record,
            param_count: None,
            params: Vec::new(),
            dispid: if constant || kind == LibraryTypeKind::Record { None } else { Some(v.memid) },
            member_flags: if v.wVarFlags.0 != 0 { Some(v.wVarFlags.0 as u32) } else { None },
            value,
        }))
    }
}

/// The integer a constant's value holds, when it holds one.
unsafe fn integer_of(v: &windows::Win32::System::Variant::VARIANT) -> Option<i64> {
    let inner = &*v.Anonymous.Anonymous;
    let data = &inner.Anonymous;
    Some(match inner.vt.0 {
        VT_I2 => data.iVal as i64,
        VT_I4 | VT_INT | VT_ERROR => data.lVal as i64,
        VT_UI1 => data.bVal as i64,
        VT_UI2 => data.uiVal as i64,
        VT_UI4 | VT_UINT => data.ulVal as i64,
        VT_I8 => data.llVal,
        VT_BOOL => data.boolVal.0 as i64,
        _ => return None,
    })
}

/// A parameter's default value as VB source text.
unsafe fn default_text(v: &windows::Win32::System::Variant::VARIANT) -> Option<String> {
    let inner = &*v.Anonymous.Anonymous;
    let data = &inner.Anonymous;
    Some(match inner.vt.0 {
        VT_BSTR => format!("\"{}\"", data.bstrVal.to_string().replace('"', "\"\"")),
        VT_BOOL => (if data.boolVal.0 != 0 { "True" } else { "False" }).to_string(),
        VT_R4 => data.fltVal.to_string(),
        VT_R8 => data.dblVal.to_string(),
        VT_EMPTY | VT_NULL => return None,
        _ => integer_of(v)?.to_string(),
    })
}
