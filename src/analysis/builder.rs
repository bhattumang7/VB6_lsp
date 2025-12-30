//! Symbol Table Builder
//!
//! Builds a symbol table by walking the tree-sitter parse tree.

use tree_sitter::{Node, Tree};
use tower_lsp::lsp_types::Url;

use super::position::{SourcePosition, SourceRange};
use super::scope::{ScopeId, ScopeKind};
use super::symbol::{ParameterInfo, SymbolId, SymbolKind, TypeInfo, Visibility};
use super::symbol_table::SymbolTable;

/// Builds a symbol table from a tree-sitter parse tree
pub struct SymbolTableBuilder<'a> {
    source: &'a str,
    table: SymbolTable,
    /// Stack of current scopes (innermost last)
    scope_stack: Vec<ScopeId>,
}

impl<'a> SymbolTableBuilder<'a> {
    /// Create a new builder
    pub fn new(uri: Url, source: &'a str) -> Self {
        let table = SymbolTable::new(uri);
        let module_scope = table.module_scope;

        Self {
            source,
            table,
            scope_stack: vec![module_scope],
        }
    }

    /// Build the symbol table from a parse tree
    pub fn build(mut self, tree: &Tree) -> SymbolTable {
        self.visit_node(&tree.root_node());
        self.table
    }

    /// Get the current scope
    fn current_scope(&self) -> ScopeId {
        *self.scope_stack.last().unwrap()
    }

    /// Push a new scope onto the stack
    fn push_scope(&mut self, kind: ScopeKind, range: SourceRange) -> ScopeId {
        let parent = Some(self.current_scope());
        let scope_id = self.table.create_scope(kind, parent, range);
        self.scope_stack.push(scope_id);
        scope_id
    }

    /// Pop the current scope
    fn pop_scope(&mut self) {
        if self.scope_stack.len() > 1 {
            self.scope_stack.pop();
        }
    }

    /// Get text content of a node
    fn node_text(&self, node: &Node) -> &str {
        node.utf8_text(self.source.as_bytes()).unwrap_or("")
    }

    /// Get range from a node
    fn node_range(&self, node: &Node) -> SourceRange {
        SourceRange::from_ts_node(node)
    }

    /// Extract visibility from a declaration node
    fn extract_visibility(&self, node: &Node) -> Visibility {
        let mut cursor = node.walk();
        for child in node.children(&mut cursor) {
            let text = self.node_text(&child).to_uppercase();
            match text.as_str() {
                "PUBLIC" => return Visibility::Public,
                "GLOBAL" => return Visibility::Global,
                "PRIVATE" => return Visibility::Private,
                "FRIEND" => return Visibility::Friend,
                _ => {}
            }
        }
        Visibility::Private // Default
    }

    /// Find a child node by field name
    fn find_field<'b>(&self, node: &'b Node<'b>, field_name: &str) -> Option<Node<'b>> {
        node.child_by_field_name(field_name)
    }

    /// Find all children of a specific kind
    fn find_children_by_kind<'b>(&self, node: &'b Node<'b>, kind: &str) -> Vec<Node<'b>> {
        let mut result = Vec::new();
        let mut cursor = node.walk();
        for child in node.children(&mut cursor) {
            if child.kind() == kind {
                result.push(child);
            }
        }
        result
    }

    /// Check if node has a specific keyword child
    fn has_child_keyword(&self, node: &Node, keyword: &str) -> bool {
        let mut cursor = node.walk();
        for child in node.children(&mut cursor) {
            if self.node_text(&child).eq_ignore_ascii_case(keyword) {
                return true;
            }
        }
        false
    }

    /// Extract type from as_clause node
    fn extract_type_from_as_clause(&self, node: &Node) -> Option<TypeInfo> {
        if let Some(type_node) = self.find_field(node, "type") {
            let name = self.node_text(&type_node).to_string();
            let is_array = name.ends_with("()") || self.find_children_by_kind(node, "array_bounds").len() > 0;
            let is_new = self.has_child_keyword(node, "new");

            Some(TypeInfo {
                name: name.trim_end_matches("()").to_string(),
                is_array,
                is_new,
            })
        } else {
            None
        }
    }

    /// Extract type from a declaration node (looks for as_clause child)
    fn extract_type(&self, node: &Node) -> Option<TypeInfo> {
        for child in self.find_children_by_kind(node, "as_clause") {
            if let Some(type_info) = self.extract_type_from_as_clause(&child) {
                return Some(type_info);
            }
        }
        None
    }

    /// Check if currently in module scope
    fn is_module_scope(&self) -> bool {
        if let Some(scope) = self.table.get_scope(self.current_scope()) {
            scope.kind == ScopeKind::Module
        } else {
            false
        }
    }

    /// Visit a node and its children
    fn visit_node(&mut self, node: &Node) {
        match node.kind() {
            // Form designer sections - SKIP entirely (these are visual control definitions, not code)
            "form_block" | "form_property_line" | "form_property_block" |
            "module_config" | "module_config_element" => {
                // Don't process form designer content as code symbols
                return;
            }

            // Declarations that create symbols
            "variable_declaration" => self.visit_variable_declaration(node),
            "constant_declaration" => self.visit_constant_declaration(node),
            "type_declaration" => self.visit_type_declaration(node),
            "enum_declaration" => self.visit_enum_declaration(node),
            "sub_declaration" => self.visit_sub_declaration(node),
            "function_declaration" => self.visit_function_declaration(node),
            "property_declaration" => self.visit_property_declaration(node),
            "declare_statement" => self.visit_declare_statement(node),
            "event_statement" => self.visit_event_statement(node),

            // Scope-creating constructs
            "with_statement" => self.visit_with_statement(node),
            "for_statement" => self.visit_for_statement(node),
            "for_each_statement" => self.visit_for_each_statement(node),

            // Labels
            "label" => self.visit_label(node),

            // Preprocessor blocks - process their children
            "preproc_if" | "preproc_elseif" | "preproc_else" => {
                self.visit_children(node);
            }

            // Default: visit children
            _ => self.visit_children(node),
        }
    }

    /// Visit all children of a node
    fn visit_children(&mut self, node: &Node) {
        let mut cursor = node.walk();
        for child in node.children(&mut cursor) {
            self.visit_node(&child);
        }
    }

    /// Visit variable declaration
    fn visit_variable_declaration(&mut self, node: &Node) {
        let visibility = self.extract_visibility(node);
        let is_local = !self.is_module_scope();

        // Find variable_list -> variable_declarator nodes
        for vl in self.find_children_by_kind(node, "variable_list") {
            for vd in self.find_children_by_kind(&vl, "variable_declarator") {
                if let Some(name_node) = self.find_field(&vd, "name") {
                    let name = self.node_text(&name_node).to_string();
                    let definition_range = self.node_range(&vd);
                    let name_range = self.node_range(&name_node);

                    let kind = if is_local {
                        SymbolKind::LocalVariable
                    } else {
                        SymbolKind::Variable
                    };

                    let symbol_id = self.table.create_symbol(
                        name,
                        kind,
                        visibility,
                        definition_range,
                        name_range,
                        self.current_scope(),
                    );

                    // Extract type info
                    if let Some(type_info) = self.extract_type(&vd) {
                        self.table.set_type_info(symbol_id, type_info);
                    }
                }
            }
        }
    }

    /// Visit constant declaration
    fn visit_constant_declaration(&mut self, node: &Node) {
        let visibility = self.extract_visibility(node);
        let is_local = !self.is_module_scope();

        for cd in self.find_children_by_kind(node, "constant_declarator") {
            if let Some(name_node) = self.find_field(&cd, "name") {
                let name = self.node_text(&name_node).to_string();
                let definition_range = self.node_range(&cd);
                let name_range = self.node_range(&name_node);

                let value = self.find_field(&cd, "value")
                    .map(|v| self.node_text(&v).to_string());

                let kind = if is_local {
                    SymbolKind::LocalConstant
                } else {
                    SymbolKind::Constant
                };

                let symbol_id = self.table.create_symbol(
                    name,
                    kind,
                    visibility,
                    definition_range,
                    name_range,
                    self.current_scope(),
                );

                if let Some(val) = value {
                    self.table.set_value(symbol_id, val);
                }

                if let Some(type_info) = self.extract_type(&cd) {
                    self.table.set_type_info(symbol_id, type_info);
                }
            }
        }
    }

    /// Visit type declaration (User-Defined Type)
    fn visit_type_declaration(&mut self, node: &Node) {
        let visibility = self.extract_visibility(node);

        if let Some(name_node) = self.find_field(node, "name") {
            let name = self.node_text(&name_node).to_string();
            let definition_range = self.node_range(node);
            let name_range = self.node_range(&name_node);

            let type_symbol_id = self.table.create_symbol(
                name,
                SymbolKind::UserDefinedType,
                visibility,
                definition_range,
                name_range,
                self.current_scope(),
            );

            // Process type members
            for tm in self.find_children_by_kind(node, "type_member") {
                if let Some(member_name_node) = self.find_field(&tm, "name") {
                    let member_name = self.node_text(&member_name_node).to_string();
                    let member_def_range = self.node_range(&tm);
                    let member_name_range = self.node_range(&member_name_node);

                    let member_id = self.table.create_symbol(
                        member_name,
                        SymbolKind::TypeMember,
                        Visibility::Public, // Type members are always public within the type
                        member_def_range,
                        member_name_range,
                        self.current_scope(),
                    );

                    if let Some(type_info) = self.extract_type(&tm) {
                        self.table.set_type_info(member_id, type_info);
                    }

                    self.table.add_member(type_symbol_id, member_id);
                }
            }
        }
    }

    /// Visit enum declaration
    fn visit_enum_declaration(&mut self, node: &Node) {
        let visibility = self.extract_visibility(node);

        if let Some(name_node) = self.find_field(node, "name") {
            let name = self.node_text(&name_node).to_string();
            let definition_range = self.node_range(node);
            let name_range = self.node_range(&name_node);

            let enum_symbol_id = self.table.create_symbol(
                name,
                SymbolKind::Enum,
                visibility,
                definition_range,
                name_range,
                self.current_scope(),
            );

            // Process enum members
            for em in self.find_children_by_kind(node, "enum_member") {
                if let Some(member_name_node) = self.find_field(&em, "name") {
                    let member_name = self.node_text(&member_name_node).to_string();
                    let member_def_range = self.node_range(&em);
                    let member_name_range = self.node_range(&member_name_node);

                    let value = self.find_field(&em, "value")
                        .map(|v| self.node_text(&v).to_string());

                    let member_id = self.table.create_symbol(
                        member_name,
                        SymbolKind::EnumMember,
                        visibility, // Enum members inherit visibility
                        member_def_range,
                        member_name_range,
                        self.current_scope(),
                    );

                    if let Some(val) = value {
                        self.table.set_value(member_id, val);
                    }

                    self.table.add_member(enum_symbol_id, member_id);
                }
            }
        }
    }

    /// Visit Sub declaration
    fn visit_sub_declaration(&mut self, node: &Node) {
        self.visit_procedure(node, SymbolKind::Sub);
    }

    /// Visit Function declaration
    fn visit_function_declaration(&mut self, node: &Node) {
        self.visit_procedure(node, SymbolKind::Function);
    }

    /// Visit Property declaration
    fn visit_property_declaration(&mut self, node: &Node) {
        let kind = if let Some(accessor) = self.find_field(node, "accessor") {
            match self.node_text(&accessor).to_uppercase().as_str() {
                "GET" => SymbolKind::PropertyGet,
                "LET" => SymbolKind::PropertyLet,
                "SET" => SymbolKind::PropertySet,
                _ => SymbolKind::PropertyGet,
            }
        } else {
            SymbolKind::PropertyGet
        };

        self.visit_procedure(node, kind);
    }

    /// Common procedure handling
    fn visit_procedure(&mut self, node: &Node, kind: SymbolKind) {
        let visibility = self.extract_visibility(node);

        if let Some(name_node) = self.find_field(node, "name") {
            let name = self.node_text(&name_node).to_string();
            let definition_range = self.node_range(node);
            let name_range = self.node_range(&name_node);

            // Create the procedure symbol
            let symbol_id = self.table.create_symbol(
                name,
                kind,
                visibility,
                definition_range,
                name_range,
                self.current_scope(),
            );

            // Extract return type for functions/property get
            if matches!(kind, SymbolKind::Function | SymbolKind::PropertyGet) {
                if let Some(type_info) = self.extract_type(node) {
                    self.table.set_type_info(symbol_id, type_info);
                }
            }

            // Create a scope for the procedure body
            let proc_scope = self.push_scope(ScopeKind::Procedure, definition_range);
            self.table.link_procedure_scope(symbol_id, proc_scope);

            // Extract and register parameters
            let parameters = self.extract_parameters(node, proc_scope);
            self.table.set_parameters(symbol_id, parameters);

            // Visit the procedure body
            for child in self.find_children_by_kind(node, "block") {
                self.visit_children(&child);
            }

            // Pop the procedure scope
            self.pop_scope();
        }
    }

    /// Extract parameters from a procedure node
    fn extract_parameters(&mut self, node: &Node, proc_scope: ScopeId) -> Vec<ParameterInfo> {
        let mut params = Vec::new();

        for pl in self.find_children_by_kind(node, "parameter_list") {
            for param in self.find_children_by_kind(&pl, "parameter") {
                if let Some(name_node) = self.find_field(&param, "name") {
                    let name = self.node_text(&name_node).to_string();
                    let param_text = self.node_text(&param).to_uppercase();

                    let by_ref = !param_text.contains("BYVAL");
                    let optional = param_text.contains("OPTIONAL");

                    let default_value = self.find_field(&param, "default")
                        .map(|v| self.node_text(&v).to_string());

                    let type_info = self.extract_type(&param);

                    let param_range = self.node_range(&param);
                    let name_range = self.node_range(&name_node);

                    // Create parameter as a symbol in procedure scope
                    let param_id = self.table.create_symbol(
                        name.clone(),
                        SymbolKind::Parameter,
                        Visibility::Private,
                        param_range,
                        name_range,
                        proc_scope,
                    );

                    if let Some(ref ti) = type_info {
                        self.table.set_type_info(param_id, ti.clone());
                    }

                    params.push(ParameterInfo {
                        name,
                        type_info,
                        by_ref,
                        optional,
                        default_value,
                        range: param_range,
                        name_range,
                    });
                }
            }
        }

        params
    }

    /// Visit Declare statement (API declaration)
    fn visit_declare_statement(&mut self, node: &Node) {
        let visibility = self.extract_visibility(node);

        if let Some(name_node) = self.find_field(node, "name") {
            let name = self.node_text(&name_node).to_string();
            let definition_range = self.node_range(node);
            let name_range = self.node_range(&name_node);

            // Determine if Sub or Function
            let kind = if self.has_child_keyword(node, "function") {
                SymbolKind::DeclareFunction
            } else {
                SymbolKind::DeclareSub
            };

            let symbol_id = self.table.create_symbol(
                name,
                kind,
                visibility,
                definition_range,
                name_range,
                self.current_scope(),
            );

            // Extract parameters (declares don't create a scope)
            let parameters = self.extract_parameters_no_scope(node);
            self.table.set_parameters(symbol_id, parameters);

            // Extract return type for functions
            if kind == SymbolKind::DeclareFunction {
                if let Some(type_info) = self.extract_type(node) {
                    self.table.set_type_info(symbol_id, type_info);
                }
            }
        }
    }

    /// Extract parameters without creating symbols (for Declare statements)
    fn extract_parameters_no_scope(&self, node: &Node) -> Vec<ParameterInfo> {
        let mut params = Vec::new();

        for pl in self.find_children_by_kind(node, "parameter_list") {
            for param in self.find_children_by_kind(&pl, "parameter") {
                if let Some(name_node) = self.find_field(&param, "name") {
                    let name = self.node_text(&name_node).to_string();
                    let param_text = self.node_text(&param).to_uppercase();

                    let by_ref = !param_text.contains("BYVAL");
                    let optional = param_text.contains("OPTIONAL");

                    let default_value = self.find_field(&param, "default")
                        .map(|v| self.node_text(&v).to_string());

                    let type_info = self.extract_type(&param);

                    params.push(ParameterInfo {
                        name,
                        type_info,
                        by_ref,
                        optional,
                        default_value,
                        range: self.node_range(&param),
                        name_range: self.node_range(&name_node),
                    });
                }
            }
        }

        params
    }

    /// Visit Event statement
    fn visit_event_statement(&mut self, node: &Node) {
        let visibility = self.extract_visibility(node);

        if let Some(name_node) = self.find_field(node, "name") {
            let name = self.node_text(&name_node).to_string();
            let definition_range = self.node_range(node);
            let name_range = self.node_range(&name_node);

            let symbol_id = self.table.create_symbol(
                name,
                SymbolKind::Event,
                visibility,
                definition_range,
                name_range,
                self.current_scope(),
            );

            let parameters = self.extract_parameters_no_scope(node);
            self.table.set_parameters(symbol_id, parameters);
        }
    }

    /// Visit With statement (creates implicit object scope)
    fn visit_with_statement(&mut self, node: &Node) {
        let range = self.node_range(node);

        // Extract the object expression
        let with_object = self.find_field(node, "object")
            .map(|obj| self.node_text(&obj).to_string());

        let scope_id = self.push_scope(ScopeKind::WithBlock, range);

        if let Some(obj) = with_object {
            self.table.set_with_object(scope_id, obj);
        }

        // Visit the block
        self.visit_children(node);

        self.pop_scope();
    }

    /// Visit For statement
    fn visit_for_statement(&mut self, node: &Node) {
        let range = self.node_range(node);

        // Create scope for loop variable
        let scope_id = self.push_scope(ScopeKind::ForLoop, range);

        // Register loop variable
        if let Some(counter_node) = self.find_field(node, "counter") {
            let name = self.node_text(&counter_node).to_string();
            let name_range = self.node_range(&counter_node);

            self.table.create_symbol(
                name,
                SymbolKind::ForLoopVariable,
                Visibility::Private,
                name_range,
                name_range,
                scope_id,
            );
        }

        // Visit the block
        self.visit_children(node);

        self.pop_scope();
    }

    /// Visit For Each statement
    fn visit_for_each_statement(&mut self, node: &Node) {
        let range = self.node_range(node);

        let scope_id = self.push_scope(ScopeKind::ForEachLoop, range);

        // Register element variable
        if let Some(element_node) = self.find_field(node, "element") {
            let name = self.node_text(&element_node).to_string();
            let name_range = self.node_range(&element_node);

            self.table.create_symbol(
                name,
                SymbolKind::ForEachVariable,
                Visibility::Private,
                name_range,
                name_range,
                scope_id,
            );
        }

        self.visit_children(node);

        self.pop_scope();
    }

    /// Visit Label
    fn visit_label(&mut self, node: &Node) {
        // Labels are the first child (identifier or integer)
        if let Some(label_node) = node.child(0) {
            let name = self.node_text(&label_node).to_string();
            let range = self.node_range(node);
            let name_range = self.node_range(&label_node);

            self.table.create_symbol(
                name,
                SymbolKind::Label,
                Visibility::Private,
                range,
                name_range,
                self.current_scope(),
            );
        }
    }
}

/// Build a symbol table from source code and tree-sitter tree
pub fn build_symbol_table(uri: Url, source: &str, tree: &Tree) -> SymbolTable {
    let builder = SymbolTableBuilder::new(uri, source);
    builder.build(tree)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::tree_sitter::TreeSitterVb6Parser;

    fn parse_and_build(source: &str) -> SymbolTable {
        let mut parser = TreeSitterVb6Parser::new().unwrap();
        let tree = parser.parse(source, None).unwrap();
        build_symbol_table(Url::parse("file:///test.bas").unwrap(), source, &tree)
    }

    #[test]
    fn test_variable_declaration() {
        let source = "Dim x As Integer";
        let table = parse_and_build(source);

        assert_eq!(table.symbol_count(), 1);
        let symbols: Vec<_> = table.all_symbols().collect();
        assert_eq!(symbols[0].name, "x");
        assert_eq!(symbols[0].kind, SymbolKind::Variable);
    }

    #[test]
    fn test_sub_declaration() {
        let source = r#"
Sub Main()
    Dim local As String
End Sub
"#;
        let table = parse_and_build(source);

        // Should have: Main (Sub) and local (LocalVariable)
        let symbols: Vec<_> = table.all_symbols().collect();
        assert!(symbols.iter().any(|s| s.name == "Main" && s.kind == SymbolKind::Sub));
        assert!(symbols.iter().any(|s| s.name == "local" && s.kind == SymbolKind::LocalVariable));
    }

    #[test]
    fn test_function_with_params() {
        let source = r#"
Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function
"#;
        let table = parse_and_build(source);

        let func: Vec<_> = table.symbols_of_kind(SymbolKind::Function).collect();
        assert_eq!(func.len(), 1);
        assert_eq!(func[0].name, "Add");
        assert_eq!(func[0].parameters.len(), 2);
        assert_eq!(func[0].parameters[0].name, "a");
        assert_eq!(func[0].parameters[1].name, "b");
    }

    #[test]
    fn test_enum_declaration() {
        let source = r#"
Public Enum Colors
    Red = 1
    Green = 2
    Blue = 3
End Enum
"#;
        let table = parse_and_build(source);

        let enums: Vec<_> = table.symbols_of_kind(SymbolKind::Enum).collect();
        assert_eq!(enums.len(), 1);
        assert_eq!(enums[0].name, "Colors");
        assert_eq!(enums[0].members.len(), 3);
    }

    #[test]
    fn test_scope_hierarchy() {
        let source = r#"
Dim moduleVar As Integer

Sub Test()
    Dim localVar As String
End Sub
"#;
        let table = parse_and_build(source);

        // moduleVar should be in module scope
        let module_var = table.lookup_symbol("moduleVar", table.module_scope);
        assert!(module_var.is_some());

        // localVar should NOT be visible from module scope
        let local_var = table.lookup_symbol("localVar", table.module_scope);
        assert!(local_var.is_none());
    }
}
