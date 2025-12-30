//! Tree-sitter grammar for VB6/VBA
//!
//! This crate provides a tree-sitter grammar for parsing Visual Basic 6
//! and VBA (Visual Basic for Applications) source code.

use tree_sitter_language::LanguageFn;

extern "C" {
    fn tree_sitter_vb6() -> *const ();
}

/// The tree-sitter [`LanguageFn`] for VB6/VBA.
///
/// # Example
///
/// ```
/// use tree_sitter::Parser;
///
/// let mut parser = Parser::new();
/// parser.set_language(&tree_sitter_vb6::LANGUAGE.into()).unwrap();
///
/// let source = "Sub Main()\nEnd Sub";
/// let tree = parser.parse(source, None).unwrap();
/// assert_eq!(tree.root_node().kind(), "source_file");
/// ```
pub const LANGUAGE: LanguageFn = unsafe { LanguageFn::from_raw(tree_sitter_vb6) };

/// Get the tree-sitter [`LanguageFn`] for VB6/VBA.
#[must_use]
pub fn language() -> LanguageFn {
    LANGUAGE
}

/// The content of the [`queries/highlights.scm`] file.
pub const HIGHLIGHTS_QUERY: &str = include_str!("../../queries/highlights.scm");

/// The content of the [`queries/locals.scm`] file.
pub const LOCALS_QUERY: &str = include_str!("../../queries/locals.scm");

/// The content of the [`queries/tags.scm`] file.
pub const TAGS_QUERY: &str = include_str!("../../queries/tags.scm");

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_can_load_grammar() {
        let mut parser = tree_sitter::Parser::new();
        parser
            .set_language(&LANGUAGE.into())
            .expect("Error loading VB6 grammar");
    }
}
