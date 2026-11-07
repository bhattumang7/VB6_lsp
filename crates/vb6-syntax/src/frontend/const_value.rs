//! Values of constant expressions (the value of a `Const`, an `Enum` member, an array bound,
//! the length of a fixed-length string, the default of an `Optional` parameter) and the
//! operators and conversions that act on them.
//!
//! A value has a type of its own (Boolean, Byte, Integer, Long, Single, Double, Currency,
//! Date, String), or is Empty, Null or Nothing. Arithmetic follows the static typing of the
//! language: the result has the wider operand's type and Overflow is raised when it does not
//! fit; an operand held in a Variant (`Const c As Variant = 5`) takes the result to the next
//! wider type instead. Converting a value to a declared type rounds halves to even, and
//! raises Overflow, Type mismatch or an invalid use of Null as the language does.

use crate::frontend::cond_compile::{format_currency, format_float, round_half_even};

/// Why a constant expression has no value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EvalError {
    /// The expression names something that is not a constant, or calls something that
    /// cannot be evaluated at compile time.
    NotConstant,
    /// A name used as a value is a Sub.
    ExpectedFunctionOrVariable,
    /// A constant is given arguments or an index.
    ExpectedArray,
    /// A member is taken of a constant (`k.x`).
    InvalidQualifier,
    /// A name qualified by an enumeration or by the module itself is not declared there.
    MemberNotFound,
    /// A `!` member is taken of a constant (`k!x`).
    QualifierMustBeCollection,
    Overflow,
    DivisionByZero,
    TypeMismatch,
    InvalidUseOfNull,
    InvalidUseOfObject,
    /// `Like` was given a pattern that is not well formed.
    InvalidPattern,
    /// `Pow` of zero to a negative power, and other calls the run time refuses.
    InvalidProcedureCall,
    /// A constant of a type that cannot hold a constant (`Const c = Nothing`).
    InvalidDataType,
    /// `Const c As Variant = Nothing`.
    ObjectNotSet,
}

impl EvalError {
    /// The compile-error code that reports it.
    pub fn code(self) -> u32 {
        use crate::frontend::diagnostics::*;
        match self {
            EvalError::NotConstant => ERR_CONSTANT_EXPRESSION_REQUIRED,
            EvalError::ExpectedFunctionOrVariable => ERR_EXPECTED_FUNCTION_OR_VARIABLE,
            EvalError::ExpectedArray => ERR_EXPECTED_ARRAY,
            EvalError::InvalidQualifier => ERR_INVALID_QUALIFIER,
            EvalError::MemberNotFound => ERR_METHOD_OR_DATA_MEMBER_NOT_FOUND,
            EvalError::QualifierMustBeCollection => ERR_QUALIFIER_MUST_BE_COLLECTION,
            EvalError::Overflow => ERR_OVERFLOW,
            EvalError::DivisionByZero => ERR_DIVISION_BY_ZERO,
            EvalError::TypeMismatch => ERR_TYPE_MISMATCH,
            EvalError::InvalidUseOfNull => ERR_INVALID_USE_OF_NULL,
            EvalError::InvalidUseOfObject => ERR_INVALID_USE_OF_OBJECT,
            EvalError::InvalidPattern => ERR_INVALID_PATTERN,
            EvalError::InvalidProcedureCall => ERR_INVALID_PROCEDURE_CALL,
            EvalError::InvalidDataType => ERR_INVALID_DATA_TYPE_FOR_CONSTANT,
            EvalError::ObjectNotSet => ERR_OBJECT_VARIABLE_NOT_SET,
        }
    }
}

pub type EvalResult = Result<ConstValue, EvalError>;

/// A constant value.
#[derive(Debug, Clone, PartialEq)]
pub enum ConstValue {
    Empty,
    Null,
    Nothing,
    Boolean(bool),
    Byte(u8),
    Integer(i16),
    Long(i32),
    Single(f32),
    Double(f64),
    /// A Currency value in ten-thousandths.
    Currency(i64),
    Date(f64),
    Str(String),
}

/// The type a constant, a bound, a length or a default value is converted to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TargetType {
    Boolean,
    Byte,
    Integer,
    Long,
    Single,
    Double,
    Currency,
    Date,
    String,
    Variant,
}

/// The type of a numeric operand, narrowest first. Currency sits beside the floating types:
/// it combines with whole numbers into Currency and with Single or Double into Double.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Debug)]
enum Rank {
    Byte,
    Integer,
    Long,
    Currency,
    Single,
    Double,
}

#[derive(Clone, Copy)]
struct Num {
    rank: Rank,
    value: f64,
    /// A Currency operand's exact ten-thousandths.
    scaled: Option<i64>,
    /// The operand is a Date: the result of `+` and `-` stays a Date.
    date: bool,
}

/// A value together with whether it is held in a Variant. An operand held in a Variant makes
/// an arithmetic result that does not fit take the next wider type instead of raising Overflow.
#[derive(Debug, Clone, PartialEq)]
pub struct Operand {
    pub value: ConstValue,
    pub variant: bool,
}

impl Operand {
    pub fn plain(value: ConstValue) -> Operand {
        Operand { value, variant: false }
    }
}

const DATE_MIN: f64 = -657_434.0;
const DATE_MAX: f64 = 2_958_465.999_999_99;
/// The Currency range in ten-thousandths is that of an `i64`.
const CURRENCY_LIMIT: f64 = 922_337_203_685_477.6;

impl ConstValue {
    /// The operand as a number. Empty and Boolean are Integer; numeric text is Double.
    fn num(&self) -> Result<Num, EvalError> {
        let n = |rank, value| Num { rank, value, scaled: None, date: false };
        match self {
            ConstValue::Empty => Ok(n(Rank::Integer, 0.0)),
            ConstValue::Boolean(b) => Ok(n(Rank::Integer, if *b { -1.0 } else { 0.0 })),
            ConstValue::Byte(v) => Ok(n(Rank::Byte, *v as f64)),
            ConstValue::Integer(v) => Ok(n(Rank::Integer, *v as f64)),
            ConstValue::Long(v) => Ok(n(Rank::Long, *v as f64)),
            ConstValue::Single(v) => Ok(n(Rank::Single, *v as f64)),
            ConstValue::Double(v) => Ok(n(Rank::Double, *v)),
            ConstValue::Date(v) => Ok(Num { rank: Rank::Double, value: *v, scaled: None, date: true }),
            ConstValue::Currency(c) => {
                Ok(Num { rank: Rank::Currency, value: *c as f64 / 10_000.0, scaled: Some(*c), date: false })
            }
            ConstValue::Str(s) => number_in_text(s).map(|t| n(Rank::Double, t.value)).ok_or(EvalError::TypeMismatch),
            ConstValue::Null => Err(EvalError::InvalidUseOfNull),
            ConstValue::Nothing => Err(EvalError::InvalidUseOfObject),
        }
    }

    /// The text of the value as `&` and `CStr` show it.
    pub fn text(&self) -> Result<String, EvalError> {
        Ok(match self {
            ConstValue::Empty => String::new(),
            ConstValue::Boolean(b) => (if *b { "True" } else { "False" }).to_string(),
            ConstValue::Byte(v) => v.to_string(),
            ConstValue::Integer(v) => v.to_string(),
            ConstValue::Long(v) => v.to_string(),
            ConstValue::Single(v) => format_float(*v as f64, 7),
            ConstValue::Double(v) => format_float(*v, 15),
            ConstValue::Date(v) if !date_in_range(*v) => return Err(EvalError::InvalidProcedureCall),
            ConstValue::Date(v) => format_date(*v),
            ConstValue::Currency(c) => format_currency(*c),
            ConstValue::Str(s) => s.clone(),
            ConstValue::Null => return Err(EvalError::InvalidUseOfNull),
            ConstValue::Nothing => return Err(EvalError::InvalidUseOfObject),
        })
    }

    /// The value rounded (halves to even) to a Long, for the bitwise operators, `\` and `Mod`.
    fn long(&self) -> Result<i32, EvalError> {
        let n = round_half_even(self.num()?.value);
        if n >= i32::MIN as f64 && n <= i32::MAX as f64 { Ok(n as i32) } else { Err(EvalError::Overflow) }
    }

    /// The value as a number of the given rank, or the error that a value of that rank
    /// cannot hold it.
    fn make(rank: Rank, value: f64) -> EvalResult {
        match rank {
            Rank::Byte => fits((0.0..=255.0).contains(&value), || ConstValue::Byte(value as u8)),
            Rank::Integer => {
                fits(value >= i16::MIN as f64 && value <= i16::MAX as f64, || ConstValue::Integer(value as i16))
            }
            Rank::Long => fits(value >= i32::MIN as f64 && value <= i32::MAX as f64, || ConstValue::Long(value as i32)),
            Rank::Single => fits(value.abs() <= f32::MAX as f64, || ConstValue::Single(value as f32)),
            Rank::Double => Ok(ConstValue::Double(value)),
            Rank::Currency => {
                fits(value.abs() <= CURRENCY_LIMIT, || ConstValue::Currency(round_half_even(value * 10_000.0) as i64))
            }
        }
    }

    /// Converts the value to a declared type, as assigning it to a variable of that type does.
    pub fn convert(&self, target: TargetType) -> EvalResult {
        match target {
            TargetType::Variant => self.convert_to_variant(),
            TargetType::String => self.convert_to_string(),
            TargetType::Boolean => self.convert_to_boolean(),
            TargetType::Date => self.convert_to_date(),
            TargetType::Byte => self.convert_to_whole(Rank::Byte),
            TargetType::Integer => self.convert_to_whole(Rank::Integer),
            TargetType::Long => Self::make(Rank::Long, round_half_even(self.num()?.value)),
            TargetType::Single => Self::make(Rank::Single, self.num()?.value),
            TargetType::Double => {
                let v = self.num()?.value;
                if v.is_finite() { Self::make(Rank::Double, v) } else { Err(EvalError::Overflow) }
            }
            TargetType::Currency => match self {
                ConstValue::Currency(c) => Ok(ConstValue::Currency(*c)),
                other => Self::make(Rank::Currency, other.num()?.value),
            },
        }
    }

    fn convert_to_variant(&self) -> EvalResult {
        match self {
            ConstValue::Nothing => Err(EvalError::ObjectNotSet),
            ConstValue::Double(d) if !d.is_finite() => Err(EvalError::Overflow),
            other => Ok(other.clone()),
        }
    }

    fn convert_to_string(&self) -> EvalResult {
        match self {
            ConstValue::Str(s) => Ok(ConstValue::Str(s.clone())),
            ConstValue::Double(d) if !d.is_finite() => Err(EvalError::Overflow),
            other => other.text().map(ConstValue::Str),
        }
    }

    fn convert_to_boolean(&self) -> EvalResult {
        match self {
            ConstValue::Boolean(b) => Ok(ConstValue::Boolean(*b)),
            ConstValue::Str(s) => match s.trim().to_ascii_lowercase().as_str() {
                "true" => Ok(ConstValue::Boolean(true)),
                "false" => Ok(ConstValue::Boolean(false)),
                _ => Ok(ConstValue::Boolean(self.num()?.value != 0.0)),
            },
            other => Ok(ConstValue::Boolean(other.num()?.value != 0.0)),
        }
    }

    fn convert_to_date(&self) -> EvalResult {
        match self {
            ConstValue::Date(d) => Ok(ConstValue::Date(*d)),
            ConstValue::Str(s) => match number_in_text(s) {
                Some(t) => date_value(t.value).map_err(|_| EvalError::TypeMismatch),
                None => parse_date(s).map(ConstValue::Date).ok_or(EvalError::TypeMismatch),
            },
            other => date_value(other.num()?.value),
        }
    }

    /// Conversion to a Byte or an Integer, which round and check the range alike.
    fn convert_to_whole(&self, rank: Rank) -> EvalResult {
        // A Boolean `True` is 255 as a Byte.
        if rank == Rank::Byte && *self == ConstValue::Boolean(true) {
            return Ok(ConstValue::Byte(255));
        }
        // A hexadecimal or octal text wider than 32 bits keeps its low 32 bits.
        if let ConstValue::Str(s) = self {
            if let Some(NumberInText { wrapped: Some(w), .. }) = number_in_text(s) {
                return Self::make(rank, w as f64);
            }
        }
        Self::make(rank, round_half_even(self.num()?.value))
    }

    /// The value as a whole number for a bound or a length: converted to a Long.
    pub fn as_long(&self) -> Result<i32, EvalError> {
        match self.convert(TargetType::Long)? {
            ConstValue::Long(v) => Ok(v),
            _ => unreachable!("conversion to Long yields a Long"),
        }
    }

    /// Whether the value is a number or numeric text (not Null, Nothing or text that is not a number).
    pub fn is_numeric_operand(&self) -> bool {
        self.num().is_ok()
    }
}

/// The value built by `make` when the number fits its type, otherwise Overflow.
fn fits(in_range: bool, value: impl FnOnce() -> ConstValue) -> EvalResult {
    if in_range { Ok(value()) } else { Err(EvalError::Overflow) }
}

/// A quotient or a power must be a finite number.
fn finite(v: ConstValue) -> EvalResult {
    match v {
        ConstValue::Double(d) if !d.is_finite() => Err(EvalError::Overflow),
        other => Ok(other),
    }
}

fn date_value(v: f64) -> EvalResult {
    if date_in_range(v) { Ok(ConstValue::Date(v)) } else { Err(EvalError::Overflow) }
}

fn date_in_range(v: f64) -> bool {
    (DATE_MIN..=DATE_MAX).contains(&v)
}

// ── Unary operators ──────────────────────────────────────────────────────────

impl Operand {
    /// Unary minus.
    pub fn negate(&self) -> Result<Operand, EvalError> {
        let keep = |value| Operand { value, variant: self.variant };
        if self.value == ConstValue::Null {
            return Ok(keep(ConstValue::Null));
        }
        let n = self.value.num()?;
        if let Some(c) = n.scaled {
            return c.checked_neg().map(|v| keep(ConstValue::Currency(v))).ok_or(EvalError::Overflow);
        }
        if matches!(self.value, ConstValue::Str(_)) {
            return ConstValue::make(Rank::Double, -n.value).map(keep);
        }
        // A Byte or a Boolean negates to an Integer; the smallest Integer, with a Variant,
        // to a Long.
        let rank = match n.rank {
            Rank::Byte => Rank::Integer,
            Rank::Integer if self.variant && n.value == i16::MIN as f64 => Rank::Long,
            other => other,
        };
        let value = ConstValue::make(rank, -n.value)?;
        Ok(keep(if n.date { ConstValue::Date(-n.value) } else { value }))
    }

    /// Unary plus.
    pub fn plus(&self) -> Result<Operand, EvalError> {
        let keep = |value| Operand { value, variant: self.variant };
        match &self.value {
            ConstValue::Null => Ok(keep(ConstValue::Null)),
            ConstValue::Str(_) => Ok(keep(ConstValue::Double(self.value.num()?.value))),
            ConstValue::Boolean(_) | ConstValue::Empty => ConstValue::make(Rank::Integer, self.value.num()?.value).map(keep),
            other => {
                other.num()?;
                Ok(keep(other.clone()))
            }
        }
    }

    /// `Not`: the bitwise complement of the operand rounded to a Long (to an Integer for an
    /// operand no wider than Integer).
    pub fn not(&self) -> Result<Operand, EvalError> {
        let keep = |value| Operand { value, variant: self.variant };
        match &self.value {
            ConstValue::Null => Ok(keep(ConstValue::Null)),
            ConstValue::Boolean(b) => Ok(keep(ConstValue::Boolean(!b))),
            other => {
                let rank = other.num()?.rank;
                let v = !other.long()?;
                Ok(keep(if rank <= Rank::Integer { ConstValue::Integer(v as i16) } else { ConstValue::Long(v) }))
            }
        }
    }
}

// ── Binary operators ─────────────────────────────────────────────────────────

/// The binary operators of the expression language.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOp {
    Mul,
    Div,
    IDiv,
    Mod,
    Add,
    Sub,
    Pow,
    Cat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
    Like,
    Is,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

/// How strings compare: by their characters, or ignoring case (`Option Compare Text`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CompareMode {
    Binary,
    Text,
}

impl Operand {
    /// Converts the value to a declared type. A number held in a Variant that lies outside
    /// the Date range is a type mismatch as a Date, not an overflow.
    pub fn convert(&self, target: TargetType) -> EvalResult {
        if target == TargetType::Date && self.variant {
            let outside = match &self.value {
                ConstValue::Date(d) => !date_in_range(*d),
                ConstValue::Str(_) => false,
                other => other.num().is_ok_and(|n| !date_in_range(n.value)),
            };
            if outside {
                return Err(EvalError::TypeMismatch);
            }
        }
        self.value.convert(target)
    }

    /// Applies a binary operator.
    pub fn binary(op: BinaryOp, l: &Operand, r: &Operand, mode: CompareMode) -> Result<Operand, EvalError> {
        let variant = l.variant || r.variant;
        let wrap = |value| Operand { value, variant };
        let (lv, rv) = (&l.value, &r.value);
        if let Some(done) = binary_shortcut(op, lv, rv) {
            return done.map(wrap);
        }
        if *lv == ConstValue::Null || *rv == ConstValue::Null {
            return null_operation(op, lv, rv).map(wrap);
        }
        binary_value(op, lv, rv, variant, mode).map(wrap)
    }
}

/// The operators whose result does not wait for the Null check: `Is`, `&`, and `And` / `Or`
/// when one operand decides.
fn binary_shortcut(op: BinaryOp, lv: &ConstValue, rv: &ConstValue) -> Option<EvalResult> {
    match op {
        BinaryOp::Is => Some(if *lv == ConstValue::Nothing && *rv == ConstValue::Nothing {
            Ok(ConstValue::Boolean(true))
        } else {
            Err(EvalError::TypeMismatch)
        }),
        BinaryOp::Cat => Some(concatenate(lv, rv)),
        BinaryOp::And | BinaryOp::Or => null_logic(op, lv, rv).transpose(),
        _ => None,
    }
}

fn concatenate(lv: &ConstValue, rv: &ConstValue) -> EvalResult {
    // Null on one side is nothing; on both sides it is Null.
    match (lv, rv) {
        (ConstValue::Null, ConstValue::Null) => Ok(ConstValue::Null),
        (ConstValue::Null, o) | (o, ConstValue::Null) => Ok(ConstValue::Str(o.text()?)),
        _ => Ok(ConstValue::Str(format!("{}{}", lv.text()?, rv.text()?))),
    }
}

/// An operation on Null yields Null (the cases handled before aside); the operand that is
/// not Null must still be usable as a value.
fn null_operation(op: BinaryOp, lv: &ConstValue, rv: &ConstValue) -> EvalResult {
    if *lv != ConstValue::Null {
        lv.value_check(op)?;
    }
    if *rv != ConstValue::Null {
        rv.value_check(op)?;
    }
    Ok(ConstValue::Null)
}

fn binary_value(op: BinaryOp, lv: &ConstValue, rv: &ConstValue, variant: bool, mode: CompareMode) -> EvalResult {
    use BinaryOp::*;
    match op {
        Add => add_values(lv, rv, variant),
        Sub => arithmetic(Arith::Sub, lv, rv, variant),
        Mul => arithmetic(Arith::Mul, lv, rv, variant),
        Div => divide_values(lv, rv),
        IDiv | Mod => integer_divide_values(op, lv, rv),
        Pow => power_values(lv, rv),
        Eq | Ne | Lt | Gt | Le | Ge => compare_values(op, lv, rv, mode),
        Like => Ok(ConstValue::Boolean(like(&lv.text()?, &rv.text()?, mode)?)),
        And | Or | Xor | Eqv | Imp => bitwise_values(op, lv, rv),
        Cat | Is => unreachable!("handled above"),
    }
}

fn add_values(lv: &ConstValue, rv: &ConstValue, variant: bool) -> EvalResult {
    if let (ConstValue::Str(a), ConstValue::Str(b)) = (lv, rv) {
        return Ok(ConstValue::Str(format!("{a}{b}")));
    }
    arithmetic(Arith::Add, lv, rv, variant)
}

fn divide_values(lv: &ConstValue, rv: &ConstValue) -> EvalResult {
    let (a, b) = (lv.num()?, rv.num()?);
    if b.value == 0.0 {
        // Zero over zero has no value to overflow to.
        return Err(if a.value == 0.0 { EvalError::Overflow } else { EvalError::DivisionByZero });
    }
    // Single with whole numbers (and Single with Single) divides to a Single.
    let single = a.rank <= Rank::Single
        && b.rank <= Rank::Single
        && a.rank.max(b.rank) == Rank::Single
        && a.rank != Rank::Currency
        && b.rank != Rank::Currency;
    let rank = if single { Rank::Single } else { Rank::Double };
    finite(ConstValue::make(rank, a.value / b.value)?)
}

fn integer_divide_values(op: BinaryOp, lv: &ConstValue, rv: &ConstValue) -> EvalResult {
    let (a, b) = (lv.long()?, rv.long()?);
    if b == 0 {
        return Err(EvalError::DivisionByZero);
    }
    let v = if op == BinaryOp::IDiv { a.checked_div(b) } else { a.checked_rem(b) }.ok_or(EvalError::Overflow)?;
    let small = lv.num()?.rank <= Rank::Integer && rv.num()?.rank <= Rank::Integer;
    Ok(if small && (-32768..=32767).contains(&v) { ConstValue::Integer(v as i16) } else { ConstValue::Long(v) })
}

fn power_values(lv: &ConstValue, rv: &ConstValue) -> EvalResult {
    let (a, b) = (lv.num()?.value, rv.num()?.value);
    if a == 0.0 && b < 0.0 {
        return Err(EvalError::InvalidProcedureCall);
    }
    let v = a.powf(b);
    if v.is_nan() {
        return Err(EvalError::InvalidProcedureCall);
    }
    finite(ConstValue::make(Rank::Double, v)?)
}

fn compare_values(op: BinaryOp, lv: &ConstValue, rv: &ConstValue, mode: CompareMode) -> EvalResult {
    use std::cmp::Ordering;
    let ord = compare(lv, rv, mode)?;
    Ok(ConstValue::Boolean(match op {
        BinaryOp::Eq => ord == Ordering::Equal,
        BinaryOp::Ne => ord != Ordering::Equal,
        BinaryOp::Lt => ord == Ordering::Less,
        BinaryOp::Gt => ord == Ordering::Greater,
        BinaryOp::Le => ord != Ordering::Greater,
        _ => ord != Ordering::Less,
    }))
}

fn bitwise_values(op: BinaryOp, lv: &ConstValue, rv: &ConstValue) -> EvalResult {
    let (a, b) = (lv.long()?, rv.long()?);
    let v = match op {
        BinaryOp::And => a & b,
        BinaryOp::Or => a | b,
        BinaryOp::Xor => a ^ b,
        BinaryOp::Eqv => !(a ^ b),
        _ => !a | b,
    };
    Ok(logic_result(lv, rv, v))
}

impl ConstValue {
    /// For an operation with a Null operand: the other operand must be a usable value.
    fn value_check(&self, op: BinaryOp) -> Result<(), EvalError> {
        match op {
            BinaryOp::Like => self.text().map(|_| ()),
            BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Gt | BinaryOp::Le | BinaryOp::Ge => match self {
                ConstValue::Str(_) => Ok(()),
                other => other.num().map(|_| ()),
            },
            // Whole-number division converts the operand to Long before the Null is seen.
            BinaryOp::IDiv | BinaryOp::Mod => self.as_long().map(|_| ()),
            _ => self.num().map(|_| ()),
        }
    }
}

/// `And` / `Or` with a Null operand: the other operand decides when it can.
fn null_logic(op: BinaryOp, l: &ConstValue, r: &ConstValue) -> Result<Option<ConstValue>, EvalError> {
    let other = match (l, r) {
        (ConstValue::Null, ConstValue::Null) => return Ok(Some(ConstValue::Null)),
        (ConstValue::Null, o) | (o, ConstValue::Null) => o,
        _ => return Ok(None),
    };
    let v = other.long()?;
    Ok(Some(match (op, v) {
        (BinaryOp::And, 0) => logic_result(other, other, 0),
        (BinaryOp::Or, v) if v != 0 => logic_result(other, other, v),
        _ => ConstValue::Null,
    }))
}

/// The type of a bitwise result: Boolean from two Booleans, Byte from two Bytes, Integer from
/// operands no wider than Integer, otherwise Long.
fn logic_result(l: &ConstValue, r: &ConstValue, v: i32) -> ConstValue {
    if matches!((l, r), (ConstValue::Boolean(_), ConstValue::Boolean(_))) {
        return ConstValue::Boolean(v != 0);
    }
    if matches!((l, r), (ConstValue::Byte(_), ConstValue::Byte(_))) {
        return ConstValue::Byte(v as u8);
    }
    let small = |x: &ConstValue| x.num().is_ok_and(|n| n.rank <= Rank::Integer);
    if small(l) && small(r) { ConstValue::Integer(v as i16) } else { ConstValue::Long(v) }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Arith {
    Add,
    Sub,
    Mul,
}

/// `+`, `-`, `*`: the result has the wider operand's type and must fit it. With a Variant
/// operand a result that does not fit takes the next wider type.
fn arithmetic(kind: Arith, l: &ConstValue, r: &ConstValue, variant: bool) -> Result<ConstValue, EvalError> {
    if let Some(done) = date_with_string(kind, l, r, variant) {
        return done;
    }
    let (a, b) = (l.num()?, r.num()?);
    if let Some(done) = date_arithmetic(kind, &a, &b, variant) {
        return done;
    }
    if let Some(done) = floating_currency_arithmetic(kind, l, r, &a, &b) {
        return done;
    }
    if let Some(done) = exact_currency_arithmetic(kind, &a, &b) {
        return done;
    }
    let rank = arithmetic_rank(l, r, &a, &b);
    let v = apply_arith(kind, a.value, b.value);
    match ConstValue::make(rank, v) {
        // The next wider type holds a Variant's result.
        Err(EvalError::Overflow) if variant => ConstValue::make(wider_rank(rank), v),
        other => other,
    }
}

fn apply_arith(kind: Arith, x: f64, y: f64) -> f64 {
    match kind {
        Arith::Add => x + y,
        Arith::Sub => x - y,
        Arith::Mul => x * y,
    }
}

/// Adding a String to a Date reads the String as a Date first.
fn date_with_string(kind: Arith, l: &ConstValue, r: &ConstValue, variant: bool) -> Option<EvalResult> {
    if kind != Arith::Add || variant {
        return None;
    }
    if !matches!((l, r), (ConstValue::Date(_), ConstValue::Str(_)) | (ConstValue::Str(_), ConstValue::Date(_))) {
        return None;
    }
    Some(day_number(l).and_then(|x| day_number(r).map(|y| x + y)).and_then(date_value))
}

fn day_number(v: &ConstValue) -> Result<f64, EvalError> {
    v.convert(TargetType::Date)?.num().map(|n| n.value)
}

/// Dates. Date - Date is a Double. With static operands a Date result must lie in the
/// range when the left operand is the Date, and is not held to it when the left operand
/// is a number or a String. With a Variant operand the result is a Date while it fits
/// the range and a plain number once it does not.
fn date_arithmetic(kind: Arith, a: &Num, b: &Num, variant: bool) -> Option<EvalResult> {
    if !(a.date || b.date) || kind == Arith::Mul {
        return None;
    }
    let v = apply_arith(kind, a.value, b.value);
    if a.date && b.date && kind == Arith::Sub {
        return Some(ConstValue::make(Rank::Double, v));
    }
    if variant {
        return Some(Ok(if date_in_range(v) { ConstValue::Date(v) } else { ConstValue::Double(v) }));
    }
    Some(if a.date { date_value(v) } else { Ok(ConstValue::Date(v)) })
}

/// Currency with Single or Double in a sum or difference: the floating operand is first
/// brought to Currency.
fn floating_currency_arithmetic(kind: Arith, l: &ConstValue, r: &ConstValue, a: &Num, b: &Num) -> Option<EvalResult> {
    let floating = |r: Rank| matches!(r, Rank::Single | Rank::Double);
    let applies = kind != Arith::Mul
        && ((a.rank == Rank::Currency && floating(b.rank) && !matches!(r, ConstValue::Str(_)))
            || (b.rank == Rank::Currency && floating(a.rank) && !matches!(l, ConstValue::Str(_))));
    if !applies {
        return None;
    }
    Some(as_currency(l).and_then(|x| as_currency(r).map(|y| (x, y))).and_then(|(x, y)| {
        let v = if kind == Arith::Add { i128::from(x) + i128::from(y) } else { i128::from(x) - i128::from(y) };
        currency_result(v)
    }))
}

/// The value converted to Currency, in ten-thousandths.
fn as_currency(v: &ConstValue) -> Result<i64, EvalError> {
    match v.convert(TargetType::Currency)? {
        ConstValue::Currency(c) => Ok(c),
        _ => unreachable!("conversion to Currency yields a Currency"),
    }
}

fn currency_result(v: i128) -> EvalResult {
    i64::try_from(v).map(ConstValue::Currency).map_err(|_| EvalError::Overflow)
}

/// Currency with whole numbers and Currency is exact.
fn exact_currency_arithmetic(kind: Arith, a: &Num, b: &Num) -> Option<EvalResult> {
    let whole = |r: Rank| matches!(r, Rank::Byte | Rank::Integer | Rank::Long | Rank::Currency);
    if !((a.rank == Rank::Currency || b.rank == Rank::Currency) && whole(a.rank) && whole(b.rank)) {
        return None;
    }
    let sa = a.scaled.map(i128::from).unwrap_or((a.value * 10_000.0) as i128);
    let sb = b.scaled.map(i128::from).unwrap_or((b.value * 10_000.0) as i128);
    let v = match kind {
        Arith::Add => sa + sb,
        Arith::Sub => sa - sb,
        Arith::Mul => {
            let p = sa * sb;
            // Round the product's last place half to even.
            let q = p.div_euclid(10_000);
            let rem = p.rem_euclid(10_000);
            if rem > 5_000 || (rem == 5_000 && q % 2 != 0) { q + 1 } else { q }
        }
    };
    Some(currency_result(v))
}

/// The type of an arithmetic result. Text operands count as Double; a Boolean or Empty as
/// Integer; two Bytes stay Byte.
fn arithmetic_rank(l: &ConstValue, r: &ConstValue, a: &Num, b: &Num) -> Rank {
    if matches!(l, ConstValue::Str(_)) || matches!(r, ConstValue::Str(_)) {
        return Rank::Double;
    }
    if a.rank == Rank::Currency || b.rank == Rank::Currency {
        // Currency with Single or Double.
        return Rank::Double;
    }
    // A Long with a Single is worked out as a Double.
    match (a.rank, b.rank) {
        (Rank::Long, Rank::Single) | (Rank::Single, Rank::Long) => Rank::Double,
        _ => a.rank.max(b.rank),
    }
}

fn wider_rank(rank: Rank) -> Rank {
    match rank {
        Rank::Byte => Rank::Integer,
        Rank::Integer => Rank::Long,
        Rank::Long | Rank::Single => Rank::Double,
        other => other,
    }
}

/// Orders two values as the comparison operators do: two strings compare as text (in the
/// given mode); a string against a number needs the string to be numeric; Empty is 0 against
/// a number and "" against text.
fn compare(l: &ConstValue, r: &ConstValue, mode: CompareMode) -> Result<std::cmp::Ordering, EvalError> {
    use std::cmp::Ordering;
    let text = |a: &str, b: &str| match mode {
        CompareMode::Binary => a.cmp(b),
        CompareMode::Text => text_key(a).cmp(&text_key(b)),
    };
    match (l, r) {
        (ConstValue::Str(a), ConstValue::Str(b)) => Ok(text(a, b)),
        (ConstValue::Empty, ConstValue::Str(b)) => Ok(text("", b)),
        (ConstValue::Str(a), ConstValue::Empty) => Ok(text(a, "")),
        (ConstValue::Empty, ConstValue::Empty) => Ok(Ordering::Equal),
        _ => compare_numbers(l, r),
    }
}

fn compare_numbers(l: &ConstValue, r: &ConstValue) -> Result<std::cmp::Ordering, EvalError> {
    let (a, b) = (l.num()?, r.num()?);
    if let (Some(x), Some(y)) = (a.scaled, b.scaled) {
        return Ok(x.cmp(&y));
    }
    // A Currency against a Single or Double: the floating operand is brought to Currency.
    let floating = |n: &Num| matches!(n.rank, Rank::Single | Rank::Double) && !n.date;
    let by_currency = (a.rank == Rank::Currency && floating(&b) && !matches!(r, ConstValue::Str(_)))
        || (b.rank == Rank::Currency && floating(&a) && !matches!(l, ConstValue::Str(_)));
    if by_currency {
        // A floating value too large for Currency is compared as the number it is.
        if let (Ok(x), Ok(y)) = (as_currency(l), as_currency(r)) {
            return Ok(x.cmp(&y));
        }
    }
    Ok(a.value.partial_cmp(&b.value).unwrap_or(std::cmp::Ordering::Equal))
}

/// The sort key of a string compared ignoring case: its letters folded to lower case.
fn text_key(s: &str) -> Vec<char> {
    s.chars().flat_map(char::to_lowercase).collect()
}

// ── Like ─────────────────────────────────────────────────────────────────────

/// `Like` pattern matching (`?`, `*`, `#`, `[list]`, `[!list]`). The pattern is read left to
/// right while the text is matched, so a malformed list (no closing bracket, or a range that runs
/// downward) is an error only when matching reaches it with a character still to match and
/// has not already decided the result. `[]` consumes nothing, and `[!]` is a list holding `!`.
fn like(text: &str, pattern: &str, mode: CompareMode) -> Result<bool, EvalError> {
    let fold = |s: &str| -> Vec<char> {
        match mode {
            CompareMode::Binary => s.chars().collect(),
            CompareMode::Text => text_key(s),
        }
    };
    like_from(&fold(text), &fold(pattern))
}

fn like_from(t: &[char], p: &[char]) -> Result<bool, EvalError> {
    let Some(&first) = p.first() else { return Ok(t.is_empty()) };
    match first {
        '*' => like_star(t, p),
        '?' => if t.is_empty() { Ok(false) } else { like_from(&t[1..], &p[1..]) },
        '#' => if t.first().is_some_and(|c| c.is_ascii_digit()) { like_from(&t[1..], &p[1..]) } else { Ok(false) },
        '[' => like_list(t, p),
        c => if t.first() == Some(&c) { like_from(&t[1..], &p[1..]) } else { Ok(false) },
    }
}

fn like_star(t: &[char], p: &[char]) -> Result<bool, EvalError> {
    for i in 0..=t.len() {
        if like_from(&t[i..], &p[1..])? {
            return Ok(true);
        }
    }
    Ok(false)
}

/// A `[list]` or `[!list]` at the start of the pattern `p`, then the rest of the pattern.
fn like_list(t: &[char], p: &[char]) -> Result<bool, EvalError> {
    if p.get(1) == Some(&']') {
        return like_from(t, &p[2..]);
    }
    let Some(&c) = t.first() else { return Ok(false) };
    let negate = p.get(1) == Some(&'!') && p.get(2) != Some(&']');
    let (found, i) = scan_list(p, c, if negate { 2 } else { 1 })?;
    if found && negate {
        return Ok(false);
    }
    let Some(rel) = p[i..].iter().position(|&x| x == ']') else {
        return Err(EvalError::InvalidPattern);
    };
    if found || negate { like_from(&t[1..], &p[i + rel + 1..]) } else { Ok(false) }
}

/// Reads the items of a list from `start` up to a match for `c` or the closing bracket; says
/// whether one matched and where reading stopped.
fn scan_list(p: &[char], c: char, start: usize) -> Result<(bool, usize), EvalError> {
    let mut i = start;
    loop {
        match p.get(i) {
            None => return Err(EvalError::InvalidPattern),
            Some(']') => return Ok((false, i)),
            Some(&lo) => {
                let (found, next) = list_item(p, i, lo, c)?;
                i = next;
                if found {
                    return Ok((true, i));
                }
            }
        }
    }
}

/// One character or one range of a list, starting at `i`: whether it holds `c`, and where the
/// next item starts.
fn list_item(p: &[char], i: usize, lo: char, c: char) -> Result<(bool, usize), EvalError> {
    let ranged = p.get(i + 1) == Some(&'-') && p.get(i + 2).is_some_and(|&h| h != ']');
    if !ranged {
        return Ok((lo == c, i + 1));
    }
    let hi = p[i + 2];
    if lo > hi {
        return Err(EvalError::InvalidPattern);
    }
    Ok((lo <= c && c <= hi, i + 3))
}

// ── Dates ────────────────────────────────────────────────────────────────────

const MONTHS: [&str; 12] = [
    "january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november",
    "december",
];

fn is_leap(y: i32) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

fn days_in_month(y: i32, m: i32) -> i32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        _ => if is_leap(y) { 29 } else { 28 },
    }
}

/// Days from 1899-12-30 (day 0) to the given date.
fn serial_of(y: i32, m: i32, d: i32) -> f64 {
    // Days from civil, then shifted to the 1899-12-30 epoch.
    let y2 = if m <= 2 { y - 1 } else { y };
    let era = if y2 >= 0 { y2 } else { y2 - 399 } / 400;
    let yoe = y2 - era * 400;
    let mp = (m + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days = era * 146_097 + doe - 719_468;
    (days + 25_569) as f64
}

fn civil_of(serial: i64) -> (i32, i32, i32) {
    let z = serial - 25_569 + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    ((if m <= 2 { y + 1 } else { y }) as i32, m as i32, d as i32)
}

/// A Date as `CStr` shows it: `m/d/yyyy`, `h:mm:ss AM`, or both.
pub fn format_date(v: f64) -> String {
    let mut whole = v.trunc();
    let mut frac = (v - whole).abs();
    if v < 0.0 {
        // A negative date keeps its time of day as the fraction past the day.
        whole = v.trunc();
    }
    let mut seconds = (frac * 86_400.0).round() as i64;
    if seconds >= 86_400 {
        seconds -= 86_400;
        whole += if v < 0.0 { -1.0 } else { 1.0 };
    }
    frac = seconds as f64;
    let mut out = String::new();
    if whole != 0.0 || seconds == 0 {
        let (y, m, d) = civil_of(whole as i64);
        out.push_str(&format!("{m}/{d}/{y}"));
    }
    if seconds != 0 {
        let h = (frac as i64) / 3600;
        let mi = ((frac as i64) % 3600) / 60;
        let s = (frac as i64) % 60;
        let (h12, ampm) = match h {
            0 => (12, "AM"),
            1..=11 => (h, "AM"),
            12 => (12, "PM"),
            _ => (h - 12, "PM"),
        };
        if !out.is_empty() {
            out.push(' ');
        }
        out.push_str(&format!("{h12}:{mi:02}:{s:02} {ampm}"));
    }
    out
}

/// Reads a date and/or time written as text: `1/2/2000`, `2000-01-02`, `Jan 2, 2000`,
/// `2 January 2000`, `1/2`, `12:30`, `12:30:15 PM`, or a date followed by a time.
pub fn parse_date(text: &str) -> Option<f64> {
    let t = text.trim();
    if t.is_empty() {
        return None;
    }
    // Split into whitespace/comma separated words; a time word contains ':' (or is AM/PM).
    let mut date_words: Vec<String> = Vec::new();
    let mut time: Option<f64> = None;
    let words: Vec<&str> = t.split(|c: char| c.is_whitespace()).filter(|w| !w.is_empty()).collect();
    let mut i = 0;
    while i < words.len() {
        let w = words[i];
        if w.contains(':') || is_meridian(w) {
            let (read, extra_words) = read_time_word(&words, i, time)?;
            time = Some(read);
            i += extra_words;
        } else {
            date_words.push(w.to_string());
        }
        i += 1;
    }
    let date = if date_words.is_empty() { None } else { Some(parse_date_part(&date_words.join(" "))?) };
    match (date, time) {
        (Some(d), Some(tm)) => Some(if d < 0.0 { d - tm } else { d + tm }),
        (Some(d), None) => Some(d),
        (None, Some(tm)) => Some(tm),
        (None, None) => None,
    }
}

fn is_pm(w: &str) -> bool {
    w.eq_ignore_ascii_case("pm") || w.eq_ignore_ascii_case("p")
}

/// Reads the time word at `words[i]` given the time read before it: the time, and how many
/// words after it were used up (an AM/PM word that follows).
fn read_time_word(words: &[&str], i: usize, time: Option<f64>) -> Option<(f64, usize)> {
    let w = words[i];
    if is_meridian(w) {
        // A bare AM/PM applies to the time word read before it.
        let base = time?;
        return Some((apply_meridian(base, is_pm(w))?, 0));
    }
    if time.is_some() {
        return None;
    }
    let (clock, meridian) = split_meridian(w);
    let v = parse_clock(clock)?;
    if let Some(pm) = meridian {
        return Some((apply_meridian(v, pm)?, 0));
    }
    match words.get(i + 1) {
        Some(next) if is_meridian(next) => Some((apply_meridian(v, is_pm(next))?, 1)),
        _ => Some((v, 0)),
    }
}

fn is_meridian(w: &str) -> bool {
    matches!(w.to_ascii_lowercase().as_str(), "am" | "pm" | "a" | "p")
}

fn split_meridian(w: &str) -> (&str, Option<bool>) {
    let lower = w.to_ascii_lowercase();
    for (suffix, pm) in [("am", false), ("pm", true)] {
        if lower.ends_with(suffix) && w.len() > 2 {
            return (&w[..w.len() - 2], Some(pm));
        }
    }
    (w, None)
}

fn apply_meridian(clock: f64, pm: bool) -> Option<f64> {
    let hours = (clock * 24.0 + 1e-9).floor();
    if !(1.0..=12.0).contains(&hours) {
        return if hours == 0.0 && !pm { Some(clock) } else { None };
    }
    let h = if hours == 12.0 { 0.0 } else { hours };
    let add = if pm { 12.0 } else { 0.0 };
    Some(clock - hours / 24.0 + (h + add) / 24.0)
}

fn parse_clock(s: &str) -> Option<f64> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() < 2 || parts.len() > 3 {
        return None;
    }
    let num = |p: &str| -> Option<i32> {
        if p.is_empty() || !p.chars().all(|c| c.is_ascii_digit()) { None } else { p.parse().ok() }
    };
    let h = num(parts[0])?;
    let m = num(parts[1])?;
    let sec = if parts.len() == 3 { num(parts[2])? } else { 0 };
    if h > 23 || m > 59 || sec > 59 {
        return None;
    }
    Some((h * 3600 + m * 60 + sec) as f64 / 86_400.0)
}

fn parse_date_part(s: &str) -> Option<f64> {
    // Numeric forms with / - . separators.
    let tokens: Vec<String> = tokenize_date(s)?;
    let (nums, month_name) = read_date_tokens(&tokens)?;
    // The written length of the year token matters for two-digit years.
    let digit_tokens: Vec<&String> = tokens.iter().filter(|t| t.chars().all(|c| c.is_ascii_digit())).collect();
    let (y, m, d) = match month_name {
        Some(m) => year_month_day_named(m, &nums, &digit_tokens)?,
        None => year_month_day_numeric(&nums, &digit_tokens)?,
    };
    if !(1..=12).contains(&m) || d < 1 || d > days_in_month(y, m) {
        return None;
    }
    let serial = serial_of(y, m, d);
    if (DATE_MIN..=DATE_MAX).contains(&serial) { Some(serial) } else { None }
}

/// The numbers of a date's tokens, and the month when one is written as a name (once at most).
fn read_date_tokens(tokens: &[String]) -> Option<(Vec<i32>, Option<i32>)> {
    let mut nums: Vec<i32> = Vec::new();
    let mut month_name: Option<i32> = None;
    for tk in tokens.iter() {
        if tk.chars().all(|c| c.is_ascii_digit()) {
            if tk.len() > 5 {
                return None;
            }
            nums.push(tk.parse().ok()?);
        } else {
            match month_of_name(tk) {
                Some(m) if month_name.is_none() => month_name = Some(m),
                _ => return None,
            }
        }
    }
    Some((nums, month_name))
}

/// The month a name or an abbreviation of at least three letters stands for.
fn month_of_name(token: &str) -> Option<i32> {
    let lower = token.to_ascii_lowercase();
    let hit = MONTHS.iter().position(|m| *m == lower || (lower.len() >= 3 && m.starts_with(lower.as_str()) && lower.len() <= m.len()));
    hit.map(|m| m as i32 + 1)
}

/// A year as written with `written_len` digits.
fn year_of(y: i32, written_len: usize) -> Option<i32> {
    if written_len <= 2 {
        // Two-digit years: 0-29 are 2000s, 30-99 the 1900s.
        Some(if y < 30 { 2000 + y } else { 1900 + y })
    } else if (100..=9999).contains(&y) {
        Some(y)
    } else {
        None
    }
}

/// Year, month and day of a date whose month is written as a name.
fn year_month_day_named(m: i32, nums: &[i32], digit_tokens: &[&String]) -> Option<(i32, i32, i32)> {
    match nums.len() {
        2 => {
            // `Jan 2 2000`, `2 Jan 2000`, `2000 Jan 2`
            let (a, b) = (nums[0], nums[1]);
            let (la, lb) = (digit_tokens[0].len(), digit_tokens[1].len());
            if la >= 3 { Some((year_of(a, la)?, m, b)) } else { Some((year_of(b, lb)?, m, a)) }
        }
        1 => Some((current_year(), m, nums[0])),
        _ => None,
    }
}

/// Year, month and day of a date written in numbers alone.
fn year_month_day_numeric(nums: &[i32], digit_tokens: &[&String]) -> Option<(i32, i32, i32)> {
    match nums.len() {
        3 => {
            let (la, lc) = (digit_tokens[0].len(), digit_tokens[2].len());
            if la >= 3 {
                // year-month-day
                Some((year_of(nums[0], la)?, nums[1], nums[2]))
            } else if nums[0] > 12 && nums[1] <= 12 {
                // day/month/year, when the first number cannot be a month
                Some((year_of(nums[2], lc)?, nums[1], nums[0]))
            } else {
                // month/day/year
                Some((year_of(nums[2], lc)?, nums[0], nums[1]))
            }
        }
        2 => Some((current_year(), nums[0], nums[1])),
        _ => None,
    }
}

/// The year it is now: a date written without a year takes it.
fn current_year() -> i32 {
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    civil_of(secs.div_euclid(86_400) + 25_569).0
}

fn tokenize_date(s: &str) -> Option<Vec<String>> {
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut kind = 0; // 1 digits, 2 letters
    for c in s.chars() {
        let k = date_char_kind(c);
        if k == 0 {
            flush_token(&mut out, &mut cur);
            if !matches!(c, '/' | '-' | '.' | ',' | ' ') {
                return None;
            }
            kind = 0;
        } else {
            if kind != 0 && kind != k {
                flush_token(&mut out, &mut cur);
            }
            kind = k;
            cur.push(c);
        }
    }
    flush_token(&mut out, &mut cur);
    Some(out)
}

/// 1 for a digit, 2 for a letter, 0 for anything else.
fn date_char_kind(c: char) -> u8 {
    if c.is_ascii_digit() {
        1
    } else if c.is_ascii_alphabetic() {
        2
    } else {
        0
    }
}

fn flush_token(out: &mut Vec<String>, cur: &mut String) {
    if !cur.is_empty() {
        out.push(std::mem::take(cur));
    }
}

// ── Built-in functions ───────────────────────────────────────────────────────

/// The functions that may be called in an `Enum` member, an array bound, a fixed-length string
/// size or an `Optional` default.
const FUNCTIONS: [&str; 16] = [
    "len", "lenb", "int", "fix", "abs", "sgn", "clng", "cint", "cbyte", "cstr", "cdbl", "csng", "ccur", "cdate",
    "cbool", "cvar",
];

/// Whether `name` is one of the functions a constant expression may call.
pub fn is_constant_function(name: &str) -> bool {
    FUNCTIONS.contains(&name.to_ascii_lowercase().as_str())
}

/// The functions the language itself defines (as opposed to those of a library): the ones the
/// constant expressions may call and those they may not.
const LANGUAGE_FUNCTIONS: [&str; 24] = [
    "len", "lenb", "int", "fix", "abs", "sgn", "clng", "cint", "cbyte", "cstr", "cdbl", "csng", "ccur", "cdate",
    "cbool", "cvar", "cdec", "cverr", "instr", "instrb", "strcomp", "string", "ubound", "lbound",
];

/// Whether `name` is a function the language itself defines.
pub fn is_language_function(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    LANGUAGE_FUNCTIONS.contains(&lower.as_str()) || lower == "date"
}

/// Calls one of those functions on one argument.
pub fn call_function(name: &str, arg: &Operand) -> Result<Operand, EvalError> {
    let lower = name.to_ascii_lowercase();
    let v = &arg.value;
    let plain = |value| Ok(Operand::plain(value));
    match lower.as_str() {
        "len" | "lenb" => length_call(&lower, v),
        "cvar" => Ok(Operand { value: v.clone(), variant: true }),
        "cstr" => plain(v.convert(TargetType::String)?),
        "cbool" => plain(v.convert(TargetType::Boolean)?),
        "cbyte" => plain(v.convert(TargetType::Byte)?),
        "cint" => plain(v.convert(TargetType::Integer)?),
        "clng" => plain(v.convert(TargetType::Long)?),
        "csng" => plain(v.convert(TargetType::Single)?),
        "cdbl" => plain(v.convert(TargetType::Double)?),
        "ccur" => plain(v.convert(TargetType::Currency)?),
        "cdate" => plain(v.convert(TargetType::Date)?),
        "abs" | "int" | "fix" | "sgn" => rounding_call(&lower, v),
        _ => unreachable!("not a constant function"),
    }
}

fn length_call(lower: &str, v: &ConstValue) -> Result<Operand, EvalError> {
    if *v == ConstValue::Null {
        return Ok(Operand { value: ConstValue::Null, variant: true });
    }
    let n = v.text()?.chars().count() as i32;
    Ok(Operand::plain(ConstValue::Long(if lower == "lenb" { n * 2 } else { n })))
}

/// `Abs`, `Int`, `Fix` and `Sgn` of one argument.
fn rounding_call(lower: &str, v: &ConstValue) -> Result<Operand, EvalError> {
    let plain = |value| Ok(Operand::plain(value));
    if *v == ConstValue::Null {
        return Ok(Operand { value: ConstValue::Null, variant: true });
    }
    let n = v.num()?;
    if lower == "sgn" {
        return plain(ConstValue::Integer(if n.value > 0.0 { 1 } else if n.value < 0.0 { -1 } else { 0 }));
    }
    let value = match lower {
        "abs" => n.value.abs(),
        "int" => n.value.floor(),
        _ => n.value.trunc(),
    };
    // The result has the argument's type; text gives a Double, a Boolean or Empty an
    // Integer, and a Date stays a Date.
    if let Some(c) = n.scaled {
        return plain(ConstValue::Currency(rounded_currency(lower, c)?));
    }
    if n.date {
        return plain(ConstValue::Date(value));
    }
    let rank = if matches!(v, ConstValue::Str(_)) { Rank::Double } else { n.rank };
    ConstValue::make(rank, value).map(Operand::plain)
}

/// `Abs`, `Int` or `Fix` of a Currency in ten-thousandths.
fn rounded_currency(lower: &str, c: i64) -> Result<i64, EvalError> {
    Ok(match lower {
        "abs" => c.checked_abs().ok_or(EvalError::Overflow)?,
        "int" => c.div_euclid(10_000) * 10_000,
        _ => (c / 10_000) * 10_000,
    })
}

// ── Numbers written as text ──────────────────────────────────────────────────

/// A number found in a string.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct NumberInText {
    pub value: f64,
    /// For a hexadecimal or octal number wider than 32 bits: its low 32 bits as a signed
    /// number, which is what an Integer or a Byte takes from it.
    pub wrapped: Option<i64>,
}

fn is_blank(c: char) -> bool {
    matches!(c, ' ' | '\t' | '\n' | '\r' | '\u{b}' | '\u{c}')
}

/// Reads the number a string holds, as converting the string to a number does: blanks around
/// it, a sign before it (blanks may follow the sign) or after it (blanks may come before the
/// sign), or brackets around it for a negative number; digits with thousands separators
/// anywhere after the first digit, a decimal point, an exponent introduced by `E` or `D`; or
/// `&H` / `&O` followed by digits alone. A hexadecimal or octal number is a signed Integer
/// from up to four hexadecimal (six octal) digits, a signed Long from up to eight (eleven),
/// and a larger number otherwise.
pub fn number_in_text(text: &str) -> Option<NumberInText> {
    let t = text.trim_matches(is_blank);
    if t.is_empty() {
        return None;
    }
    if let Some(rest) = t.strip_prefix('&') {
        return radix_in_text(rest);
    }
    // Brackets give a negative number; otherwise a sign leads or trails.
    let (negative, body) = if let Some(inner) = t.strip_prefix('(').and_then(|r| r.strip_suffix(')')) {
        (true, inner.trim_matches(is_blank))
    } else if let Some(r) = t.strip_prefix(['+', '-']) {
        (t.starts_with('-'), r.trim_start_matches(is_blank))
    } else if let Some(r) = t.strip_suffix(['+', '-']) {
        (t.ends_with('-'), r.trim_end_matches(is_blank))
    } else {
        (false, t)
    };
    let value = decimal_in_text(body)?;
    Some(NumberInText { value: if negative { -value } else { value }, wrapped: None })
}

/// The number after the `&` of `&H` or `&O` text.
fn radix_in_text(rest: &str) -> Option<NumberInText> {
    let mut chars = rest.chars();
    let radix = match chars.next()? {
        'h' | 'H' => 16,
        'o' | 'O' => 8,
        _ => return None,
    };
    let digits = chars.as_str();
    if digits.is_empty() || !digits.chars().all(|c| c.is_digit(radix)) {
        return None;
    }
    let mut v: u128 = 0;
    for c in digits.chars() {
        v = (v * radix as u128 + c.to_digit(radix).unwrap() as u128).min(u128::MAX >> 8);
    }
    let significant = digits.trim_start_matches('0').len();
    let wrapped = (v > 0xFFFF_FFFF).then(|| {
        let low = (v & 0xFFFF_FFFF) as u32;
        low as i32 as i64
    });
    Some(NumberInText { value: radix_value(v, radix, significant), wrapped })
}

/// The signed value of `&H`/`&O` digits: 16 bits when they fit in the short form, else 32 bits,
/// else the unwrapped magnitude.
fn radix_value(v: u128, radix: u32, significant: usize) -> f64 {
    let (short_limit, digits_short): (u128, usize) =
        if radix == 16 { (0xFFFF, 4) } else { (0o177777, 6) };
    if significant <= digits_short && v <= short_limit {
        sign_wrapped(v, 0x8000, 65_536.0)
    } else if v <= 0xFFFF_FFFF {
        sign_wrapped(v, 0x8000_0000, 4_294_967_296.0)
    } else {
        v as f64
    }
}

/// `v`, or `v - modulus` when `v` is at or above the sign `threshold`.
fn sign_wrapped(v: u128, threshold: u128, modulus: f64) -> f64 {
    if v >= threshold { v as f64 - modulus } else { v as f64 }
}

/// Digits with separators, a point and an exponent; no sign, no blanks.
fn decimal_in_text(body: &str) -> Option<f64> {
    let mut clean = String::new();
    let mut chars = body.chars().peekable();
    if !read_mantissa(&mut chars, &mut clean) {
        return None;
    }
    if chars.peek().is_some() {
        read_exponent(&mut chars, &mut clean)?;
    }
    clean.parse::<f64>().ok()
}

/// Reads digits, thousands separators after the first digit and one point into `clean`; says
/// whether a digit was found.
fn read_mantissa(chars: &mut std::iter::Peekable<std::str::Chars<'_>>, clean: &mut String) -> bool {
    let mut seen_digit = false;
    let mut seen_point = false;
    while let Some(&c) = chars.peek() {
        match c {
            '0'..='9' => {
                seen_digit = true;
                clean.push(c);
            }
            ',' if seen_digit => {}
            '.' if !seen_point => {
                seen_point = true;
                clean.push(c);
            }
            _ => break,
        }
        chars.next();
    }
    seen_digit
}

/// Reads an exponent (`E` or `D`, an optional sign, then digits alone) that must end the text.
fn read_exponent(chars: &mut std::iter::Peekable<std::str::Chars<'_>>, clean: &mut String) -> Option<()> {
    if !matches!(chars.next()?, 'e' | 'E' | 'd' | 'D') {
        return None;
    }
    clean.push('e');
    if let Some(&sign) = chars.peek() {
        if sign == '+' || sign == '-' {
            clean.push(sign);
            chars.next();
        }
    }
    let mut exponent_digits = 0;
    for c in chars.by_ref() {
        if !c.is_ascii_digit() {
            return None;
        }
        exponent_digits += 1;
        clean.push(c);
    }
    if exponent_digits == 0 {
        return None;
    }
    Some(())
}
