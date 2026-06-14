//! Hover-text / signature formatting from bound declarations.
//!
//! Identifier names are rendered from the **source text** at each declaration's
//! `name_span`, not from the scanner's interned symbol. The interner
//! canonicalizes an identifier to the casing of its first occurrence in the
//! module (VBA-style), so a parameter written `b` could intern as `B` if `B`
//! appeared earlier. Slicing the source preserves exactly what the user wrote.
//!
//! Type names (for `As <UserType>`) are still taken from the scanner, since a
//! type *use* does not carry the declaration's span.

use crate::frontend::ast::{ProcKind, Span};
use crate::frontend::scanner::ScannerContext;
use crate::sema::symbol::{BoundEnumDecl, BoundEnumMember, BoundProc, BoundTypeDecl, BoundVar};
use crate::sema::types::VbaType;

/// Slice the original source spelling at a name span (Windows-1252 → text).
/// Returns an empty string for an unrecorded (dummy) span.
pub fn name_at(src: &[u8], span: Span) -> String {
    if span.len == 0 {
        return String::new();
    }
    let lo = span.start as usize;
    let hi = (span.start + span.len) as usize;
    src.get(lo..hi)
        .map(|bytes| bytes.iter().map(|&b| b as char).collect())
        .unwrap_or_default()
}

/// Render a [`VbaType`] as VB6 type text (e.g. `Long`, `TPoint`, `String()`).
pub fn type_str(ctx: &ScannerContext, t: &VbaType) -> String {
    match t {
        VbaType::Variant => "Variant".into(),
        VbaType::Integer => "Integer".into(),
        VbaType::Long => "Long".into(),
        VbaType::Single => "Single".into(),
        VbaType::Double => "Double".into(),
        VbaType::Currency => "Currency".into(),
        VbaType::Date => "Date".into(),
        VbaType::String => "String".into(),
        VbaType::Object => "Object".into(),
        VbaType::Boolean => "Boolean".into(),
        VbaType::Decimal => "Decimal".into(),
        VbaType::Byte => "Byte".into(),
        VbaType::UserDefined(sym) => {
            let n = if *sym == 0 { String::new() } else { ctx.symbol(*sym as usize).name.clone() };
            if n.is_empty() { "Object".into() } else { n }
        }
        VbaType::Array(inner) => format!("{}()", type_str(ctx, inner)),
    }
}

fn proc_keyword(kind: ProcKind) -> &'static str {
    match kind {
        ProcKind::Sub => "Sub",
        ProcKind::Function => "Function",
        ProcKind::PropGet => "Property Get",
        ProcKind::PropLet => "Property Let",
        ProcKind::PropSet => "Property Set",
    }
}

/// Render a procedure signature, e.g.
/// `Public Function Add(a As Long, b As Long) As Long`.
pub fn proc_signature(ctx: &ScannerContext, src: &[u8], p: &BoundProc) -> String {
    let vis = if p.is_public { "Public " } else { "Private " };
    let kw = proc_keyword(p.kind);
    let name = name_at(src, p.name_span);

    let params: Vec<String> = p
        .params
        .iter()
        .map(|param| {
            let mut s = String::new();
            if param.flags.optional {
                s.push_str("Optional ");
            }
            if param.flags.param_array {
                s.push_str("ParamArray ");
            } else if param.flags.by_val {
                s.push_str("ByVal ");
            } else if param.flags.by_ref {
                s.push_str("ByRef ");
            }
            s.push_str(&name_at(src, param.name_span));
            if param.flags.is_array {
                s.push_str("()");
            }
            s.push_str(" As ");
            s.push_str(&type_str(ctx, &param.vba_type));
            s
        })
        .collect();

    let ret = match p.kind {
        ProcKind::Function | ProcKind::PropGet => {
            format!(" As {}", type_str(ctx, &p.ret_type))
        }
        _ => String::new(),
    };

    format!("{vis}{kw} {name}({}){ret}", params.join(", "))
}

/// Render a variable/constant declaration, e.g. `Public gCount As Long` or
/// `Const K As Long`.
pub fn var_signature(ctx: &ScannerContext, src: &[u8], v: &BoundVar) -> String {
    let name = name_at(src, v.name_span);
    let ty = type_str(ctx, &v.vba_type);
    if v.is_const {
        format!("Const {name} As {ty}")
    } else {
        let vis = if v.is_public { "Public" } else { "Private" };
        format!("{vis} {name} As {ty}")
    }
}

/// Render a parameter, e.g. `x As Long`.
pub fn param_signature(ctx: &ScannerContext, src: &[u8], name_span: Span, ty: &VbaType) -> String {
    format!("{} As {}", name_at(src, name_span), type_str(ctx, ty))
}

/// Render a `Type` declaration header, e.g. `Type TPoint`.
pub fn type_decl_signature(src: &[u8], t: &BoundTypeDecl) -> String {
    format!("Type {}", name_at(src, t.name_span))
}

/// Render an `Enum` declaration header, e.g. `Enum EColor`.
pub fn enum_decl_signature(src: &[u8], e: &BoundEnumDecl) -> String {
    format!("Enum {}", name_at(src, e.name_span))
}

/// Render an enum member, e.g. `EColor.Red = 0`.
pub fn enum_member_signature(src: &[u8], e: &BoundEnumDecl, m: &BoundEnumMember) -> String {
    format!("{}.{} = {}", name_at(src, e.name_span), name_at(src, m.name_span), m.value)
}
