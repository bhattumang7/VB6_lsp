//! VBA type system: semantic types produced by the binder.

/// A VB6 semantic type.
///
/// In VBA, every value is a Variant unless an explicit `As <type>` clause
/// narrows it.  This enum tracks the declared (static) type; untyped
/// expressions carry `Variant`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum VbaType {
    Variant,
    Integer,
    Long,
    Single,
    Double,
    Currency,
    Date,
    String,
    Object,
    Boolean,
    Decimal,
    Byte,
    /// User-defined type name (sym_id into the scanner's symbol table).
    UserDefined(u32),
    /// Array of some element type (e.g. `String()`, `Long()`).
    Array(Box<VbaType>),
}

impl VbaType {
    /// Convert from an AST `BuiltinType { kind }` constant.
    ///
    /// Kind values use the 5-bit type-kind encoding carried by
    /// `ExprNode::BuiltinType` and `ExprNode::TypeSpec`.
    pub fn from_kind(kind: u32) -> Self {
        match kind {
            2  => VbaType::Integer,
            3  => VbaType::Long,
            4  => VbaType::Single,
            5  => VbaType::Double,
            6  => VbaType::Currency,
            7  => VbaType::Date,
            8  => VbaType::String,
            9  => VbaType::Object,
            11 => VbaType::Boolean,
            12 => VbaType::Variant,
            14 => VbaType::Decimal,
            17 => VbaType::Byte,
            _  => VbaType::Variant,
        }
    }
}

impl Default for VbaType {
    fn default() -> Self {
        VbaType::Variant
    }
}
