//! Symbol Definitions
//!
//! Defines the Symbol struct and related types for the symbol table.

use super::position::SourceRange;
use super::scope::ScopeId;

/// Unique identifier for a symbol within the symbol table
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SymbolId(pub u32);

/// The kind of symbol
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SymbolKind {
    // Module-level declarations
    Variable,
    Constant,
    UserDefinedType,
    Enum,
    EnumMember,
    TypeMember,

    // Procedures
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
    Event,
    DeclareFunction,
    DeclareSub,

    // Local scope
    Parameter,
    LocalVariable,
    LocalConstant,

    // Special
    ForLoopVariable,
    ForEachVariable,
    Label,

    // Form controls (TextBox, Label, Button, etc.)
    FormControl,
}

impl SymbolKind {
    /// Get the LSP SymbolKind equivalent
    pub fn to_lsp(&self) -> tower_lsp::lsp_types::SymbolKind {
        use tower_lsp::lsp_types::SymbolKind as LspKind;
        match self {
            SymbolKind::Variable
            | SymbolKind::LocalVariable
            | SymbolKind::ForLoopVariable
            | SymbolKind::ForEachVariable => LspKind::VARIABLE,
            SymbolKind::Constant | SymbolKind::LocalConstant => LspKind::CONSTANT,
            SymbolKind::UserDefinedType => LspKind::STRUCT,
            SymbolKind::Enum => LspKind::ENUM,
            SymbolKind::EnumMember => LspKind::ENUM_MEMBER,
            SymbolKind::TypeMember => LspKind::FIELD,
            SymbolKind::Sub | SymbolKind::DeclareSub => LspKind::FUNCTION,
            SymbolKind::Function | SymbolKind::DeclareFunction => LspKind::FUNCTION,
            SymbolKind::PropertyGet | SymbolKind::PropertyLet | SymbolKind::PropertySet => {
                LspKind::PROPERTY
            }
            SymbolKind::Event => LspKind::EVENT,
            SymbolKind::Parameter => LspKind::VARIABLE,
            SymbolKind::Label => LspKind::NULL,
            SymbolKind::FormControl => LspKind::FIELD,
        }
    }

    /// Get completion item kind
    pub fn to_completion_kind(&self) -> tower_lsp::lsp_types::CompletionItemKind {
        use tower_lsp::lsp_types::CompletionItemKind;
        match self {
            SymbolKind::Variable
            | SymbolKind::LocalVariable
            | SymbolKind::Parameter
            | SymbolKind::ForLoopVariable
            | SymbolKind::ForEachVariable => CompletionItemKind::VARIABLE,
            SymbolKind::Constant | SymbolKind::LocalConstant => CompletionItemKind::CONSTANT,
            SymbolKind::Sub | SymbolKind::DeclareSub => CompletionItemKind::FUNCTION,
            SymbolKind::Function | SymbolKind::DeclareFunction => CompletionItemKind::FUNCTION,
            SymbolKind::PropertyGet | SymbolKind::PropertyLet | SymbolKind::PropertySet => {
                CompletionItemKind::PROPERTY
            }
            SymbolKind::UserDefinedType => CompletionItemKind::STRUCT,
            SymbolKind::Enum => CompletionItemKind::ENUM,
            SymbolKind::EnumMember => CompletionItemKind::ENUM_MEMBER,
            SymbolKind::TypeMember => CompletionItemKind::FIELD,
            SymbolKind::Event => CompletionItemKind::EVENT,
            SymbolKind::Label => CompletionItemKind::REFERENCE,
            SymbolKind::FormControl => CompletionItemKind::FIELD,
        }
    }

    /// Check if this symbol kind creates a scope
    pub fn creates_scope(&self) -> bool {
        matches!(
            self,
            SymbolKind::Sub
                | SymbolKind::Function
                | SymbolKind::PropertyGet
                | SymbolKind::PropertyLet
                | SymbolKind::PropertySet
        )
    }

    /// Check if this is a procedure-like symbol
    pub fn is_procedure(&self) -> bool {
        matches!(
            self,
            SymbolKind::Sub
                | SymbolKind::Function
                | SymbolKind::PropertyGet
                | SymbolKind::PropertyLet
                | SymbolKind::PropertySet
                | SymbolKind::DeclareFunction
                | SymbolKind::DeclareSub
        )
    }

    /// Check if this is a callable symbol
    pub fn is_callable(&self) -> bool {
        self.is_procedure() || matches!(self, SymbolKind::Event)
    }

    /// Get a display string for the kind
    pub fn display_name(&self) -> &'static str {
        match self {
            SymbolKind::Variable | SymbolKind::LocalVariable => "Variable",
            SymbolKind::Constant | SymbolKind::LocalConstant => "Constant",
            SymbolKind::UserDefinedType => "Type",
            SymbolKind::Enum => "Enum",
            SymbolKind::EnumMember => "Enum Member",
            SymbolKind::TypeMember => "Field",
            SymbolKind::Sub | SymbolKind::DeclareSub => "Sub",
            SymbolKind::Function | SymbolKind::DeclareFunction => "Function",
            SymbolKind::PropertyGet => "Property Get",
            SymbolKind::PropertyLet => "Property Let",
            SymbolKind::PropertySet => "Property Set",
            SymbolKind::Event => "Event",
            SymbolKind::Parameter => "Parameter",
            SymbolKind::ForLoopVariable | SymbolKind::ForEachVariable => "Loop Variable",
            SymbolKind::Label => "Label",
            SymbolKind::FormControl => "Control",
        }
    }
}

/// Visibility level for symbols
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Visibility {
    #[default]
    Private,
    Public,
    Friend,
    Global,
}

impl Visibility {
    /// Get display string
    pub fn as_str(&self) -> &'static str {
        match self {
            Visibility::Public => "Public",
            Visibility::Private => "Private",
            Visibility::Friend => "Friend",
            Visibility::Global => "Global",
        }
    }
}

/// Type information for a symbol
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TypeInfo {
    /// The type name (e.g., "Integer", "String", "MyClass")
    pub name: String,
    /// Whether this is an array type
    pub is_array: bool,
    /// Whether this is a New expression type (for classes)
    pub is_new: bool,
}

impl TypeInfo {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            is_array: false,
            is_new: false,
        }
    }

    pub fn array(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            is_array: true,
            is_new: false,
        }
    }

    /// Format for display (e.g., "Integer()" for arrays)
    pub fn display(&self) -> String {
        if self.is_array {
            format!("{}()", self.name)
        } else {
            self.name.clone()
        }
    }
}

/// Parameter information for procedures
#[derive(Debug, Clone)]
pub struct ParameterInfo {
    /// Parameter name
    pub name: String,
    /// Parameter type
    pub type_info: Option<TypeInfo>,
    /// Whether passed by reference (default in VB6)
    pub by_ref: bool,
    /// Whether optional
    pub optional: bool,
    /// Default value expression (for optional params)
    pub default_value: Option<String>,
    /// Position range of the entire parameter declaration
    pub range: SourceRange,
    /// Position range of just the name
    pub name_range: SourceRange,
}

impl ParameterInfo {
    /// Format parameter for signature display
    pub fn format_signature(&self) -> String {
        let mut parts = Vec::new();

        if self.optional {
            parts.push("Optional".to_string());
        }

        if self.by_ref {
            parts.push("ByRef".to_string());
        } else {
            parts.push("ByVal".to_string());
        }

        parts.push(self.name.clone());

        if let Some(ref ti) = self.type_info {
            parts.push(format!("As {}", ti.display()));
        }

        if let Some(ref default) = self.default_value {
            parts.push(format!("= {}", default));
        }

        parts.join(" ")
    }
}

/// A symbol definition in the symbol table
#[derive(Debug, Clone)]
pub struct Symbol {
    /// Unique identifier
    pub id: SymbolId,
    /// Symbol name (case-preserved, but lookups are case-insensitive)
    pub name: String,
    /// Symbol kind
    pub kind: SymbolKind,
    /// Visibility
    pub visibility: Visibility,
    /// Type information (if applicable)
    pub type_info: Option<TypeInfo>,
    /// The range of the entire declaration
    pub definition_range: SourceRange,
    /// The range of just the name (for precise go-to-definition)
    pub name_range: SourceRange,
    /// The scope this symbol belongs to
    pub scope_id: ScopeId,
    /// For procedures: parameters
    pub parameters: Vec<ParameterInfo>,
    /// For enums/types: member symbol IDs
    pub members: Vec<SymbolId>,
    /// Documentation/comments associated with this symbol
    pub documentation: Option<String>,
    /// Value (for constants and enum members)
    pub value: Option<String>,
}

impl Symbol {
    /// Create a new symbol
    pub fn new(
        id: SymbolId,
        name: String,
        kind: SymbolKind,
        visibility: Visibility,
        definition_range: SourceRange,
        name_range: SourceRange,
        scope_id: ScopeId,
    ) -> Self {
        Self {
            id,
            name,
            kind,
            visibility,
            type_info: None,
            definition_range,
            name_range,
            scope_id,
            parameters: Vec::new(),
            members: Vec::new(),
            documentation: None,
            value: None,
        }
    }

    /// Format the symbol as a signature for hover display
    pub fn format_signature(&self) -> String {
        match self.kind {
            SymbolKind::Sub | SymbolKind::DeclareSub => {
                let params = self.format_parameters();
                format!("{} Sub {}({})", self.visibility.as_str(), self.name, params)
            }
            SymbolKind::Function | SymbolKind::DeclareFunction => {
                let params = self.format_parameters();
                let ret_type = self
                    .type_info
                    .as_ref()
                    .map(|t| t.display())
                    .unwrap_or_else(|| "Variant".to_string());
                format!(
                    "{} Function {}({}) As {}",
                    self.visibility.as_str(),
                    self.name,
                    params,
                    ret_type
                )
            }
            SymbolKind::PropertyGet => {
                let params = self.format_parameters();
                let ret_type = self
                    .type_info
                    .as_ref()
                    .map(|t| t.display())
                    .unwrap_or_else(|| "Variant".to_string());
                format!(
                    "{} Property Get {}({}) As {}",
                    self.visibility.as_str(),
                    self.name,
                    params,
                    ret_type
                )
            }
            SymbolKind::PropertyLet => {
                let params = self.format_parameters();
                format!(
                    "{} Property Let {}({})",
                    self.visibility.as_str(),
                    self.name,
                    params
                )
            }
            SymbolKind::PropertySet => {
                let params = self.format_parameters();
                format!(
                    "{} Property Set {}({})",
                    self.visibility.as_str(),
                    self.name,
                    params
                )
            }
            SymbolKind::Variable | SymbolKind::LocalVariable => {
                let type_str = self
                    .type_info
                    .as_ref()
                    .map(|t| t.display())
                    .unwrap_or_else(|| "Variant".to_string());
                format!("{} {} As {}", self.visibility.as_str(), self.name, type_str)
            }
            SymbolKind::Constant | SymbolKind::LocalConstant => {
                let value = self.value.as_deref().unwrap_or("?");
                format!(
                    "{} Const {} = {}",
                    self.visibility.as_str(),
                    self.name,
                    value
                )
            }
            SymbolKind::Parameter => {
                let type_str = self
                    .type_info
                    .as_ref()
                    .map(|t| t.display())
                    .unwrap_or_else(|| "Variant".to_string());
                format!("Parameter {} As {}", self.name, type_str)
            }
            SymbolKind::Event => {
                let params = self.format_parameters();
                format!("{} Event {}({})", self.visibility.as_str(), self.name, params)
            }
            SymbolKind::UserDefinedType => {
                format!("{} Type {}", self.visibility.as_str(), self.name)
            }
            SymbolKind::Enum => {
                format!("{} Enum {}", self.visibility.as_str(), self.name)
            }
            SymbolKind::EnumMember => {
                if let Some(ref val) = self.value {
                    format!("{} = {}", self.name, val)
                } else {
                    self.name.clone()
                }
            }
            SymbolKind::TypeMember => {
                let type_str = self
                    .type_info
                    .as_ref()
                    .map(|t| t.display())
                    .unwrap_or_else(|| "Variant".to_string());
                format!("{} As {}", self.name, type_str)
            }
            SymbolKind::ForLoopVariable | SymbolKind::ForEachVariable => {
                format!("(loop variable) {}", self.name)
            }
            SymbolKind::Label => {
                format!("{}:", self.name)
            }
            SymbolKind::FormControl => {
                let type_str = self
                    .type_info
                    .as_ref()
                    .map(|t| t.display())
                    .unwrap_or_else(|| "Control".to_string());
                format!("{} As {}", self.name, type_str)
            }
        }
    }

    fn format_parameters(&self) -> String {
        self.parameters
            .iter()
            .map(|p| p.format_signature())
            .collect::<Vec<_>>()
            .join(", ")
    }
}
