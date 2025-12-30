//! AST Converter
//!
//! Converts tree-sitter parse trees into the existing Vb6Ast structure.
//! This maintains compatibility with the existing LSP implementation.

use tree_sitter::{Node, Tree};
use super::ast::*;

/// Convert a tree-sitter tree to a Vb6Ast
pub fn convert_tree(tree: &Tree, source: &str) -> Vb6Ast {
    let mut ast = Vb6Ast::new();
    let root = tree.root_node();

    // Walk all top-level children
    let mut cursor = root.walk();
    for child in root.children(&mut cursor) {
        convert_node(&child, source, &mut ast);
    }

    ast
}

/// Convert a single node and its relevant children
fn convert_node(node: &Node, source: &str, ast: &mut Vb6Ast) {
    match node.kind() {
        "option_statement" => convert_option(node, source, ast),
        "attribute_statement" => convert_attribute(node, source, ast),
        "variable_declaration" => convert_variable(node, source, ast),
        "constant_declaration" => convert_constant(node, source, ast),
        "type_declaration" => convert_type(node, source, ast),
        "enum_declaration" => convert_enum(node, source, ast),
        "declare_statement" => convert_declare(node, source, ast),
        "event_statement" => convert_event(node, source, ast),
        "implements_statement" => convert_implements(node, source, ast),
        "deftype_statement" => convert_deftype(node, source, ast),
        "sub_declaration" => convert_sub(node, source, ast),
        "function_declaration" => convert_function(node, source, ast),
        "property_declaration" => convert_property(node, source, ast),
        "preproc_const" => convert_preproc_const(node, source, ast),
        "preproc_if" | "preproc_elseif" | "preproc_else" => convert_preproc_if(node, source, ast),
        "comment" => convert_comment(node, source, ast),
        _ => {}
    }
}

/// Convert preprocessor constant
fn convert_preproc_const(node: &Node, source: &str, ast: &mut Vb6Ast) {
    // #Const is similar to a regular constant, store it as a constant with special visibility
    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();
        let value = find_field(node, "value")
            .map(|v| node_text(&v, source).to_string())
            .unwrap_or_default();

        ast.add_constant(Constant {
            name: format!("#Const {}", name),  // Prefix to indicate preprocessor constant
            value,
            visibility: Visibility::Private,
            line: node_line(node),
        });
    }
}

/// Convert preprocessor if block - recursively process children
fn convert_preproc_if(node: &Node, source: &str, ast: &mut Vb6Ast) {
    // Process all children nodes within the preprocessor block
    let mut cursor = node.walk();
    for child in node.children(&mut cursor) {
        match child.kind() {
            // Recursively process module elements inside preprocessor blocks
            "variable_declaration" => convert_variable(&child, source, ast),
            "constant_declaration" => convert_constant(&child, source, ast),
            "type_declaration" => convert_type(&child, source, ast),
            "enum_declaration" => convert_enum(&child, source, ast),
            "declare_statement" => convert_declare(&child, source, ast),
            "event_statement" => convert_event(&child, source, ast),
            "implements_statement" => convert_implements(&child, source, ast),
            "deftype_statement" => convert_deftype(&child, source, ast),
            "sub_declaration" => convert_sub(&child, source, ast),
            "function_declaration" => convert_function(&child, source, ast),
            "property_declaration" => convert_property(&child, source, ast),
            "preproc_elseif" | "preproc_else" => convert_preproc_if(&child, source, ast),
            "comment" => convert_comment(&child, source, ast),
            _ => {}
        }
    }
}

/// Get text content of a node
fn node_text<'a>(node: &Node, source: &'a str) -> &'a str {
    node.utf8_text(source.as_bytes()).unwrap_or("")
}

/// Get the line number (0-indexed)
fn node_line(node: &Node) -> usize {
    node.start_position().row
}

/// Convert Option statement
fn convert_option(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let text = node_text(node, source);
    ast.add_option(node_line(node), text);
}

/// Convert Attribute statement
fn convert_attribute(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let text = node_text(node, source);
    ast.add_attribute(node_line(node), text);
}

/// Convert comment
fn convert_comment(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let text = node_text(node, source);
    ast.add_comment(node_line(node), text);
}

/// Convert Declare statement (external API declaration)
fn convert_declare(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();
        let parameters = convert_parameters(node, source);

        // Check if it's a Sub or Function by looking for return type
        let proc_type = if find_children_by_kind(node, "as_clause").is_empty() {
            ProcedureType::Sub
        } else {
            ProcedureType::Function
        };

        let return_type = find_children_by_kind(node, "as_clause")
            .first()
            .and_then(|ac| extract_type_from_as_clause(ac, source));

        // Declare statements don't have end lines
        ast.add_procedure(Procedure {
            name: format!("Declare {}", name),  // Prefix to indicate it's a Declare
            proc_type,
            visibility,
            line,
            parameters,
            return_type,
            end_line: None,
        });
    }
}

/// Convert Event statement
fn convert_event(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();
        let parameters = convert_parameters(node, source);

        // Events are similar to Sub declarations but without bodies
        ast.add_procedure(Procedure {
            name: format!("Event {}", name),  // Prefix to indicate it's an Event
            proc_type: ProcedureType::Sub,
            visibility,
            line,
            parameters,
            return_type: None,
            end_line: None,
        });
    }
}

/// Convert Implements statement
fn convert_implements(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let text = node_text(node, source);
    // Store as an attribute since it's a module-level directive
    ast.add_attribute(node_line(node), text);
}

/// Convert DefType statement
fn convert_deftype(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let text = node_text(node, source);
    // Store as an option since it affects compilation behavior
    ast.add_option(node_line(node), text);
}

/// Extract visibility from a declaration node
fn extract_visibility(node: &Node, source: &str) -> Visibility {
    let mut cursor = node.walk();
    for child in node.children(&mut cursor) {
        let text = node_text(&child, source).to_uppercase();
        match text.as_str() {
            "PUBLIC" | "GLOBAL" => return Visibility::Public,
            "PRIVATE" => return Visibility::Private,
            "FRIEND" => return Visibility::Friend,
            _ => {}
        }
    }
    Visibility::Private // Default
}

/// Find a child node by field name
fn find_field<'a>(node: &'a Node, field_name: &str) -> Option<Node<'a>> {
    node.child_by_field_name(field_name)
}

/// Find child nodes by kind
fn find_children_by_kind<'a>(node: &'a Node, kind: &str) -> Vec<Node<'a>> {
    let mut result = Vec::new();
    let mut cursor = node.walk();
    for child in node.children(&mut cursor) {
        if child.kind() == kind {
            result.push(child);
        }
    }
    result
}

/// Extract type from as_clause node
fn extract_type_from_as_clause(node: &Node, source: &str) -> Option<String> {
    if let Some(type_node) = find_field(node, "type") {
        Some(node_text(&type_node, source).to_string())
    } else {
        None
    }
}

/// Convert variable declaration
fn convert_variable(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    // Find variable_list -> variable_declarator nodes
    for vl in find_children_by_kind(node, "variable_list") {
        for vd in find_children_by_kind(&vl, "variable_declarator") {
            if let Some(name_node) = find_field(&vd, "name") {
                let name = node_text(&name_node, source).to_string();

                // Check for array bounds
                let is_array = find_children_by_kind(&vd, "array_bounds").len() > 0;

                // Get type from as_clause
                let var_type = find_children_by_kind(&vd, "as_clause")
                    .first()
                    .and_then(|ac| extract_type_from_as_clause(ac, source));

                ast.add_variable(Variable {
                    name,
                    var_type,
                    visibility,
                    line,
                    is_array,
                });
            }
        }
    }
}

/// Convert constant declaration
fn convert_constant(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    // Find constant_declarator nodes
    for cd in find_children_by_kind(node, "constant_declarator") {
        if let Some(name_node) = find_field(&cd, "name") {
            let name = node_text(&name_node, source).to_string();

            let value = find_field(&cd, "value")
                .map(|v| node_text(&v, source).to_string())
                .unwrap_or_default();

            ast.add_constant(Constant {
                name,
                value,
                visibility,
                line,
            });
        }
    }
}

/// Convert type declaration
fn convert_type(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();

        // Extract type members
        let mut members = Vec::new();
        for tm in find_children_by_kind(node, "type_member") {
            if let Some(member_name_node) = find_field(&tm, "name") {
                let member_name = node_text(&member_name_node, source).to_string();

                let member_type = find_children_by_kind(&tm, "as_clause")
                    .first()
                    .and_then(|ac| extract_type_from_as_clause(ac, source))
                    .unwrap_or_else(|| "Variant".to_string());

                members.push(TypeMember {
                    name: member_name,
                    member_type,
                });
            }
        }

        ast.add_user_type(UserType {
            name,
            visibility,
            line,
            members,
        });
    }
}

/// Convert enum declaration
fn convert_enum(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();

        // Extract enum members
        let mut members = Vec::new();
        for em in find_children_by_kind(node, "enum_member") {
            if let Some(member_name_node) = find_field(&em, "name") {
                let member_name = node_text(&member_name_node, source).to_string();

                let value = find_field(&em, "value")
                    .and_then(|v| node_text(&v, source).parse::<i32>().ok());

                members.push(EnumMember {
                    name: member_name,
                    value,
                });
            }
        }

        ast.add_enum(Enumeration {
            name,
            visibility,
            line,
            members,
        });
    }
}

/// Convert parameters from a parameter_list node
fn convert_parameters(node: &Node, source: &str) -> Vec<Parameter> {
    let mut params = Vec::new();

    for pl in find_children_by_kind(node, "parameter_list") {
        for param in find_children_by_kind(&pl, "parameter") {
            if let Some(name_node) = find_field(&param, "name") {
                let name = node_text(&name_node, source).to_string();
                let param_text = node_text(&param, source).to_uppercase();

                let by_ref = !param_text.contains("BYVAL");
                let optional = param_text.contains("OPTIONAL");

                let param_type = find_children_by_kind(&param, "as_clause")
                    .first()
                    .and_then(|ac| extract_type_from_as_clause(ac, source));

                params.push(Parameter {
                    name,
                    param_type,
                    by_ref,
                    optional,
                });
            }
        }
    }

    params
}

/// Find the end line of a procedure (End Sub/Function/Property)
fn find_end_line(node: &Node) -> Option<usize> {
    Some(node.end_position().row)
}

/// Convert Sub declaration
fn convert_sub(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();
        let parameters = convert_parameters(node, source);
        let end_line = find_end_line(node);

        ast.add_procedure(Procedure {
            name,
            proc_type: ProcedureType::Sub,
            visibility,
            line,
            parameters,
            return_type: None,
            end_line,
        });
    }
}

/// Convert Function declaration
fn convert_function(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();
        let parameters = convert_parameters(node, source);
        let end_line = find_end_line(node);

        // Get return type from as_clause after parameters
        let return_type = find_children_by_kind(node, "as_clause")
            .first()
            .and_then(|ac| extract_type_from_as_clause(ac, source));

        ast.add_procedure(Procedure {
            name,
            proc_type: ProcedureType::Function,
            visibility,
            line,
            parameters,
            return_type,
            end_line,
        });
    }
}

/// Convert Property declaration
fn convert_property(node: &Node, source: &str, ast: &mut Vb6Ast) {
    let visibility = extract_visibility(node, source);
    let line = node_line(node);

    // Determine property type from accessor field
    let proc_type = if let Some(accessor) = find_field(node, "accessor") {
        let accessor_text = node_text(&accessor, source).to_uppercase();
        match accessor_text.as_str() {
            "GET" => ProcedureType::PropertyGet,
            "LET" => ProcedureType::PropertyLet,
            "SET" => ProcedureType::PropertySet,
            _ => ProcedureType::PropertyGet,
        }
    } else {
        ProcedureType::PropertyGet
    };

    if let Some(name_node) = find_field(node, "name") {
        let name = node_text(&name_node, source).to_string();
        let parameters = convert_parameters(node, source);
        let end_line = find_end_line(node);

        // Get return type for Property Get
        let return_type = if proc_type == ProcedureType::PropertyGet {
            find_children_by_kind(node, "as_clause")
                .first()
                .and_then(|ac| extract_type_from_as_clause(ac, source))
        } else {
            None
        };

        ast.add_procedure(Procedure {
            name,
            proc_type,
            visibility,
            line,
            parameters,
            return_type,
            end_line,
        });
    }
}

/// Extract parse errors from the tree
pub fn extract_errors(tree: &Tree, source: &str) -> Vec<ParseErrorInfo> {
    let mut errors = Vec::new();
    collect_errors(&tree.root_node(), source, &mut errors);
    errors
}

/// Recursively collect error nodes
fn collect_errors(node: &Node, source: &str, errors: &mut Vec<ParseErrorInfo>) {
    if node.is_error() || node.is_missing() {
        errors.push(ParseErrorInfo {
            message: if node.is_missing() {
                format!("Missing: {}", node.kind())
            } else {
                format!("Syntax error at: {}", node_text(node, source))
            },
            line: node.start_position().row,
            column: node.start_position().column,
            end_line: node.end_position().row,
            end_column: node.end_position().column,
        });
    }

    let mut cursor = node.walk();
    for child in node.children(&mut cursor) {
        collect_errors(&child, source, errors);
    }
}

/// Parse error information
#[derive(Debug, Clone)]
pub struct ParseErrorInfo {
    pub message: String,
    pub line: usize,
    pub column: usize,
    pub end_line: usize,
    pub end_column: usize,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::tree_sitter::TreeSitterVb6Parser;

    fn parse_and_convert(source: &str) -> Vb6Ast {
        let mut parser = TreeSitterVb6Parser::new().unwrap();
        let tree = parser.parse(source, None).unwrap();
        convert_tree(&tree, source)
    }

    #[test]
    fn test_convert_variable() {
        let source = "Dim x As Integer";
        let ast = parse_and_convert(source);
        assert_eq!(ast.variables.len(), 1);
        assert_eq!(ast.variables[0].name, "x");
        assert_eq!(ast.variables[0].var_type, Some("Integer".to_string()));
    }

    #[test]
    fn test_convert_sub() {
        let source = r#"
Sub Main()
    x = 10
End Sub
"#;
        let ast = parse_and_convert(source);
        assert_eq!(ast.procedures.len(), 1);
        assert_eq!(ast.procedures[0].name, "Main");
        assert_eq!(ast.procedures[0].proc_type, ProcedureType::Sub);
    }

    #[test]
    fn test_convert_function() {
        let source = r#"
Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function
"#;
        let ast = parse_and_convert(source);
        assert_eq!(ast.procedures.len(), 1);
        assert_eq!(ast.procedures[0].name, "Add");
        assert_eq!(ast.procedures[0].proc_type, ProcedureType::Function);
        assert_eq!(ast.procedures[0].parameters.len(), 2);
        assert_eq!(ast.procedures[0].return_type, Some("Integer".to_string()));
    }
}
