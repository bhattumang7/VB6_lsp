//! Tree-sitter wrapper for VB6 parsing
//!
//! Provides incremental parsing capabilities using the tree-sitter-vb6 grammar.

use tree_sitter::{Parser, Tree, Language, Query, QueryCursor};
use streaming_iterator::StreamingIterator;

/// Tree-sitter parser wrapper for VB6
pub struct TreeSitterVb6Parser {
    parser: Parser,
}

impl TreeSitterVb6Parser {
    /// Create a new tree-sitter VB6 parser
    pub fn new() -> Result<Self, String> {
        let mut parser = Parser::new();
        let language: Language = tree_sitter_vb6::LANGUAGE.into();

        parser.set_language(&language)
            .map_err(|e| format!("Error loading VB6 grammar: {}", e))?;

        Ok(Self { parser })
    }

    /// Parse VB6 source code, optionally using a previous tree for incremental parsing
    pub fn parse(&mut self, source: &str, old_tree: Option<&Tree>) -> Option<Tree> {
        self.parser.parse(source, old_tree)
    }

    /// Get the tree-sitter language for queries
    pub fn language(&self) -> Language {
        tree_sitter_vb6::LANGUAGE.into()
    }
}

impl Default for TreeSitterVb6Parser {
    fn default() -> Self {
        Self::new().expect("Failed to create VB6 parser")
    }
}

/// Query helper for tree-sitter queries
pub struct VB6QueryRunner {
    language: Language,
}

impl VB6QueryRunner {
    pub fn new() -> Self {
        Self {
            language: tree_sitter_vb6::LANGUAGE.into(),
        }
    }

    /// Create a query from a pattern string
    pub fn create_query(&self, pattern: &str) -> Result<Query, tree_sitter::QueryError> {
        Query::new(&self.language, pattern)
    }

    /// Run a query on a tree and return matches
    pub fn run_query<'a>(
        &self,
        query: &'a Query,
        tree: &'a Tree,
        source: &'a [u8],
    ) -> Vec<QueryMatch<'a>> {
        let mut cursor = QueryCursor::new();
        let root = tree.root_node();

        let mut results = Vec::new();
        let mut matches = cursor.matches(query, root, source);

        while let Some(m) = matches.next() {
            results.push(QueryMatch {
                pattern_index: m.pattern_index,
                captures: m.captures.iter().map(|c| {
                    QueryCapture {
                        name: query.capture_names()[c.index as usize].to_string(),
                        node: c.node,
                        text: c.node.utf8_text(source).unwrap_or("").to_string(),
                    }
                }).collect(),
            });
        }

        results
    }
}

impl Default for VB6QueryRunner {
    fn default() -> Self {
        Self::new()
    }
}

/// A query match result
pub struct QueryMatch<'a> {
    pub pattern_index: usize,
    pub captures: Vec<QueryCapture<'a>>,
}

/// A captured node from a query
pub struct QueryCapture<'a> {
    pub name: String,
    pub node: tree_sitter::Node<'a>,
    pub text: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parser_creation() {
        let parser = TreeSitterVb6Parser::new();
        assert!(parser.is_ok());
    }

    #[test]
    fn test_basic_parse() {
        let mut parser = TreeSitterVb6Parser::new().unwrap();
        let source = r#"
Option Explicit

Dim x As Integer

Sub Main()
    x = 10
End Sub
"#;
        let tree = parser.parse(source, None);
        assert!(tree.is_some());

        let tree = tree.unwrap();
        let root = tree.root_node();
        assert_eq!(root.kind(), "source_file");
    }
}
