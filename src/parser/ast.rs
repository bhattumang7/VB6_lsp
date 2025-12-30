//! Abstract Syntax Tree definitions for VB6

use std::collections::HashMap;

/// Complete VB6 AST for a source file
#[derive(Debug, Clone)]
pub struct Vb6Ast {
    pub file_type: FileType,
    pub options: Vec<String>,
    pub attributes: Vec<String>,
    pub comments: HashMap<usize, String>,
    pub imports: Vec<String>,
    pub variables: Vec<Variable>,
    pub constants: Vec<Constant>,
    pub user_types: Vec<UserType>,
    pub enums: Vec<Enumeration>,
    pub procedures: Vec<Procedure>,
    pub statements: HashMap<usize, String>,
}

impl Vb6Ast {
    pub fn new() -> Self {
        Self {
            file_type: FileType::Module,
            options: Vec::new(),
            attributes: Vec::new(),
            comments: HashMap::new(),
            imports: Vec::new(),
            variables: Vec::new(),
            constants: Vec::new(),
            user_types: Vec::new(),
            enums: Vec::new(),
            procedures: Vec::new(),
            statements: HashMap::new(),
        }
    }

    pub fn add_option(&mut self, _line: usize, content: &str) {
        self.options.push(content.to_string());
    }

    pub fn add_attribute(&mut self, _line: usize, content: &str) {
        self.attributes.push(content.to_string());
    }

    pub fn add_comment(&mut self, line: usize, content: &str) {
        self.comments.insert(line, content.to_string());
    }

    pub fn add_variable(&mut self, var: Variable) {
        self.variables.push(var);
    }

    pub fn add_constant(&mut self, constant: Constant) {
        self.constants.push(constant);
    }

    pub fn add_user_type(&mut self, user_type: UserType) {
        self.user_types.push(user_type);
    }

    pub fn add_enum(&mut self, enumeration: Enumeration) {
        self.enums.push(enumeration);
    }

    pub fn add_procedure(&mut self, proc: Procedure) {
        self.procedures.push(proc);
    }

    pub fn add_statement(&mut self, line: usize, content: &str) {
        self.statements.insert(line, content.to_string());
    }
}

impl Default for Vb6Ast {
    fn default() -> Self {
        Self::new()
    }
}

/// VB6 file type
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FileType {
    Module,      // .bas
    Class,       // .cls
    Form,        // .frm
    UserControl, // .ctl
}

/// Visibility modifier
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    Public,
    Private,
    Friend,
}

/// Variable declaration
#[derive(Debug, Clone)]
pub struct Variable {
    pub name: String,
    pub var_type: Option<String>,
    pub visibility: Visibility,
    pub line: usize,
    pub is_array: bool,
}

/// Constant declaration
#[derive(Debug, Clone)]
pub struct Constant {
    pub name: String,
    pub value: String,
    pub visibility: Visibility,
    pub line: usize,
}

/// User-defined Type
#[derive(Debug, Clone)]
pub struct UserType {
    pub name: String,
    pub visibility: Visibility,
    pub line: usize,
    pub members: Vec<TypeMember>,
}

/// Type member
#[derive(Debug, Clone)]
pub struct TypeMember {
    pub name: String,
    pub member_type: String,
}

/// Enumeration
#[derive(Debug, Clone)]
pub struct Enumeration {
    pub name: String,
    pub visibility: Visibility,
    pub line: usize,
    pub members: Vec<EnumMember>,
}

/// Enum member
#[derive(Debug, Clone)]
pub struct EnumMember {
    pub name: String,
    pub value: Option<i32>,
}

/// Procedure type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcedureType {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

/// Procedure (Sub/Function/Property)
#[derive(Debug, Clone)]
pub struct Procedure {
    pub name: String,
    pub proc_type: ProcedureType,
    pub visibility: Visibility,
    pub line: usize,
    pub parameters: Vec<Parameter>,
    pub return_type: Option<String>,
    pub end_line: Option<usize>,
}

/// Parameter
#[derive(Debug, Clone)]
pub struct Parameter {
    pub name: String,
    pub param_type: Option<String>,
    pub by_ref: bool,
    pub optional: bool,
}

/// Symbol information for LSP operations
#[derive(Debug, Clone)]
pub struct Symbol {
    pub name: String,
    pub kind: SymbolKind,
    pub line: usize,
    pub column: usize,
    pub documentation: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SymbolKind {
    Variable,
    Constant,
    Type,
    Enum,
    Function,
    Sub,
    Property,
    Parameter,
    Field,
}
