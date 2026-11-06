//! The names a module can be given with `Attribute VB_Name`, and the text of an attribute's value.
//!
//! Setting the name of a module can fail in four ways. The failures are decided by the name alone
//! (an empty name, a bracket in it, `Main`, a character a standard module's name cannot hold) or
//! by the project around the module (the name of the project, of a library it references, of
//! another module). When a module sets its name more than once, the failures of the earlier
//! settings are replaced by the later ones, except that a [`NameFailure::Reserved`]
//! failure (an empty name, a bracket, the project's name or a library's) stays.

use crate::frontend::diagnostics::{
    Diagnostic, ERR_MAIN_NOT_ALLOWED_AS_NAME, ERR_NAME_CONFLICTS, ERR_NOT_LEGAL_OBJECT_NAME,
};
use crate::frontend::parser::ModuleKind;
use crate::frontend::token::Span;

/// The value of an attribute's string once its backslash escapes are applied: `\n` is a line
/// feed, `\r` a carriage return, and a backslash before any other character (a backslash
/// included) is dropped and leaves that character. A backslash that ends the string is dropped.
pub fn unescape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut chars = text.chars();
    while let Some(c) = chars.next() {
        if c != '\\' {
            out.push(c);
            continue;
        }
        match chars.next() {
            Some('n') => out.push('\n'),
            Some('r') => out.push('\r'),
            Some(other) => out.push(other),
            None => {}
        }
    }
    out
}

/// The longest name a standard module can have.
pub const MAX_STANDARD_MODULE_NAME_LEN: usize = 127;

/// Whether `name` is a legal name for a standard module: a letter, then letters, digits and
/// underscores, at most [`MAX_STANDARD_MODULE_NAME_LEN`] of them. A character outside ASCII
/// counts as a letter.
pub fn is_legal_standard_module_name(name: &str) -> bool {
    let mut chars = name.chars();
    let Some(first) = chars.next() else { return false };
    (first.is_ascii_alphabetic() || !first.is_ascii())
        && chars.all(|c| c.is_ascii_alphanumeric() || c == '_' || !c.is_ascii())
        && name.chars().count() <= MAX_STANDARD_MODULE_NAME_LEN
}

/// Why a module's name was refused.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NameFailure {
    /// A standard module's name that is not a legal name.
    Illegal,
    /// `Main`, in any letter case.
    Main,
    /// The name is empty, has a bracket in it, or is the name of the project or of a library:
    /// the module keeps the failure whatever name it is given afterwards.
    Reserved,
    /// The name is another module's.
    TakenByModule,
}

/// What a module's name is checked against besides itself.
#[derive(Clone, Copy, Debug)]
pub struct NameScope<'a> {
    pub kind: ModuleKind,
    /// The project's name.
    pub project: Option<&'a str>,
    /// The names of the libraries the project references.
    pub libraries: &'a [&'a str],
    /// The names of the project's other modules.
    pub modules: &'a [&'a str],
}

impl NameScope<'_> {
    /// Why `name` is refused for the module, if it is.
    pub fn failure(&self, name: &str) -> Option<NameFailure> {
        let same = |other: &str| other.eq_ignore_ascii_case(name);
        if name.is_empty() || name.contains(['[', ']']) {
            return Some(NameFailure::Reserved);
        }
        // A form, a UserControl and the other designers take their name from the designer
        // section; the name set here is only compared with it.
        let designer = self.kind == ModuleKind::Form;
        if self.kind == ModuleKind::Standard && !is_legal_standard_module_name(name) {
            return Some(NameFailure::Illegal);
        }
        if !designer && same("Main") {
            return Some(NameFailure::Main);
        }
        if self.project.is_some_and(same) || self.libraries.iter().any(|l| same(l)) {
            return Some(NameFailure::Reserved);
        }
        if !designer && self.modules.iter().any(|m| same(m)) {
            return Some(NameFailure::TakenByModule);
        }
        None
    }
}

/// The names a module is given, in order, and the failure the last of them leaves.
#[derive(Debug, Default)]
pub struct ModuleNameSetting {
    failed: Option<(NameFailure, String, Span)>,
}

impl ModuleNameSetting {
    pub fn new() -> Self {
        Self::default()
    }

    /// The module is given `name`, written at `at`.
    pub fn set(&mut self, scope: &NameScope, name: &str, at: Span) {
        if matches!(self.failed, Some((NameFailure::Reserved, ..))) {
            return;
        }
        self.failed = scope.failure(name).map(|f| (f, name.to_string(), at));
    }

    /// The diagnostic for the failure the settings left, if any.
    pub fn diagnostic(&self) -> Option<Diagnostic> {
        let (failure, name, at) = self.failed.as_ref()?;
        Some(match failure {
            NameFailure::Illegal => Diagnostic::with_subject(ERR_NOT_LEGAL_OBJECT_NAME, *at, name.clone()),
            NameFailure::Main => Diagnostic::new(ERR_MAIN_NOT_ALLOWED_AS_NAME, *at),
            NameFailure::Reserved | NameFailure::TakenByModule => Diagnostic::new(ERR_NAME_CONFLICTS, *at),
        })
    }
}
