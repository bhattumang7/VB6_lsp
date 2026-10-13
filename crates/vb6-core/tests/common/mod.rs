//! Helpers shared by the member-resolution tests.
#![allow(dead_code)]

use vb6_core::frontend::diagnostics::Diagnostic;
use vb6_core::session::{Location, Session, SessionConfig};

pub fn session(files: &[(&str, &str)]) -> Session {
    session_with(files, SessionConfig::default())
}

pub fn session_with(files: &[(&str, &str)], config: SessionConfig) -> Session {
    Session::with_config(
        files.iter().map(|(p, s)| (p.to_string(), s.as_bytes().to_vec())).collect(),
        config,
    )
}

/// Byte offset of the first character of `member` inside the first occurrence
/// of `context` in `src`.
pub fn at(src: &str, context: &str, member: &str) -> u32 {
    let c = src.find(context).unwrap_or_else(|| panic!("`{context}` not in source"));
    let m = context.find(member).unwrap_or_else(|| panic!("`{member}` not in `{context}`"));
    (c + m) as u32
}

/// Byte offset of the `n`th (0-based) occurrence of `needle` in `src`.
pub fn nth(src: &str, needle: &str, n: usize) -> u32 {
    src.match_indices(needle).nth(n).unwrap_or_else(|| panic!("`{needle}` #{n} not in source")).0 as u32
}

pub fn text<'a>(src: &'a str, loc: Location) -> &'a str {
    &src[loc.span.start as usize..(loc.span.start + loc.span.len) as usize]
}

/// `(module, start offset, spelling)` of each location.
pub fn sites(srcs: &[&str], locs: &[Location]) -> Vec<(usize, u32, String)> {
    locs.iter().map(|l| (l.module, l.span.start, text(srcs[l.module], *l).to_string())).collect()
}

/// The diagnostics of `module` with the given code, as `(start, spelling)`.
pub fn coded(s: &Session, src: &str, module: usize, code: u32) -> Vec<(u32, String)> {
    s.diagnostics(module)
        .into_iter()
        .filter(|d: &Diagnostic| d.code == code)
        .map(|d| {
            (d.span.start, src[d.span.start as usize..(d.span.start + d.span.len) as usize].to_string())
        })
        .collect()
}
