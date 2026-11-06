//! Values of conditional-compilation expressions (`#Const`, `#If`) and the
//! project's conditional-compilation arguments.
//!
//! A conditional-compilation value is a typed Variant: Empty, Null, Boolean, Integer,
//! Long, Single, Double, Currency, Date, String or Nothing. Operators coerce as VB does
//! for constants: arithmetic keeps the narrowest type that holds both operands and
//! raises Overflow when the result does not fit it, comparisons yield Boolean, and
//! `And`/`Or`/`Not`/... act on the bits of the operands rounded to Long. An operation
//! that VB rejects yields the error number it raises.

/// The error numbers an evaluation raises.
pub const ERR_OVERFLOW: u32 = 6;
pub const ERR_INVALID_PROCEDURE_CALL: u32 = 5;
pub const ERR_DIVISION_BY_ZERO: u32 = 11;
pub const ERR_TYPE_MISMATCH: u32 = 13;
pub const ERR_INVALID_USE_OF_NULL: u32 = 94;
/// A call of something other than a built-in function (`Variable not defined`).
pub const ERR_VARIABLE_NOT_DEFINED: u32 = 0x9caf;

const FUNCTIONS: [&str; 17] = [
    "len", "lenb", "abs", "int", "fix", "sgn", "cbool", "cbyte", "ccur", "cdate", "cdbl", "cint", "clng", "csng", "cstr",
    "cvar", "cverr",
];

/// The result of evaluating: a value, or the error number it raised.
pub type CcResult = Result<CcValue, u32>;

/// A conditional-compilation value.
#[derive(Debug, Clone, PartialEq)]
pub enum CcValue {
    Empty,
    Null,
    Nothing,
    Boolean(bool),
    Integer(i16),
    Long(i32),
    Single(f32),
    Double(f64),
    /// A Currency value in ten-thousandths.
    Currency(i64),
    Date(f64),
    Str(String),
    /// An error value (`CVErr`).
    Error(i32),
}

/// Operators of the expression language, in the evaluator's tagging.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CcOp {
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

/// A numeric operand: its rank (Integer < Long < Single < Double, with Currency apart)
/// and its value.
#[derive(Clone, Copy, PartialEq, PartialOrd)]
enum Rank {
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
    /// Set for a Currency operand: its exact ten-thousandths.
    scaled: Option<i64>,
}

impl CcValue {
    pub const TRUE: CcValue = CcValue::Boolean(true);
    pub const FALSE: CcValue = CcValue::Boolean(false);

    fn from_bool(b: bool) -> CcValue {
        CcValue::Boolean(b)
    }

    /// The operand as a number: Empty and Boolean are Integer, numeric text is Double.
    fn num(&self) -> Result<Num, u32> {
        let n = |rank, value| Num { rank, value, scaled: None };
        match self {
            CcValue::Empty => Ok(n(Rank::Integer, 0.0)),
            CcValue::Boolean(b) => Ok(n(Rank::Integer, if *b { -1.0 } else { 0.0 })),
            CcValue::Integer(v) => Ok(n(Rank::Integer, *v as f64)),
            CcValue::Long(v) => Ok(n(Rank::Long, *v as f64)),
            CcValue::Single(v) => Ok(n(Rank::Single, *v as f64)),
            CcValue::Double(v) | CcValue::Date(v) => Ok(n(Rank::Double, *v)),
            CcValue::Currency(c) => Ok(Num { rank: Rank::Currency, value: *c as f64 / 10_000.0, scaled: Some(*c) }),
            CcValue::Str(s) => parse_number(s).map(|v| n(Rank::Double, v)).ok_or(ERR_TYPE_MISMATCH),
            CcValue::Null => Err(ERR_INVALID_USE_OF_NULL),
            CcValue::Nothing | CcValue::Error(_) => Err(ERR_TYPE_MISMATCH),
        }
    }

    /// The text of the value, as `&` sees it.
    pub fn text(&self) -> Result<String, u32> {
        Ok(match self {
            CcValue::Empty => String::new(),
            CcValue::Boolean(b) => (if *b { "True" } else { "False" }).to_string(),
            CcValue::Integer(v) => v.to_string(),
            CcValue::Long(v) => v.to_string(),
            CcValue::Single(v) => format_float(*v as f64, 7),
            CcValue::Double(v) | CcValue::Date(v) => format_float(*v, 15),
            CcValue::Currency(c) => format_currency(*c),
            CcValue::Str(s) => s.clone(),
            CcValue::Null | CcValue::Nothing | CcValue::Error(_) => return Err(ERR_TYPE_MISMATCH),
        })
    }

    /// The value rounded (halves to even) to a Long, for the bitwise operators, `\` and
    /// `Mod`; Overflow when it does not fit.
    fn long(&self) -> Result<i32, u32> {
        let n = round_half_even(self.num()?.value);
        if n >= i32::MIN as f64 && n <= i32::MAX as f64 { Ok(n as i32) } else { Err(ERR_OVERFLOW) }
    }

    /// Whether a condition with this value is taken: a non-zero number.
    pub fn is_true(&self) -> Result<bool, u32> {
        Ok(self.num()?.value != 0.0)
    }

    /// A numeric result of the given rank, or Overflow if it does not fit.
    fn make(rank: Rank, value: f64) -> CcResult {
        match rank {
            Rank::Integer => {
                if value >= i16::MIN as f64 && value <= i16::MAX as f64 { Ok(CcValue::Integer(value as i16)) } else { Err(ERR_OVERFLOW) }
            }
            Rank::Long => {
                if value >= i32::MIN as f64 && value <= i32::MAX as f64 { Ok(CcValue::Long(value as i32)) } else { Err(ERR_OVERFLOW) }
            }
            Rank::Single => {
                if value.abs() <= f32::MAX as f64 { Ok(CcValue::Single(value as f32)) } else { Err(ERR_OVERFLOW) }
            }
            Rank::Double => Ok(CcValue::Double(value)),
            Rank::Currency => {
                let scaled = (value * 10_000.0).round();
                if scaled.abs() <= i64::MAX as f64 { Ok(CcValue::Currency(scaled as i64)) } else { Err(ERR_OVERFLOW) }
            }
        }
    }

    /// Unary minus.
    pub fn negate(&self) -> CcResult {
        if *self == CcValue::Null {
            return Ok(CcValue::Null);
        }
        let n = self.num()?;
        match n.scaled {
            Some(c) => c.checked_neg().map(CcValue::Currency).ok_or(ERR_OVERFLOW),
            None => {
                // Negating the smallest Integer widens to Long.
                let rank = if n.rank == Rank::Integer && n.value == i16::MIN as f64 { Rank::Long } else { n.rank };
                CcValue::make(rank, -n.value)
            }
        }
    }

    /// Unary plus.
    pub fn plus(&self) -> CcResult {
        match self {
            CcValue::Null => Ok(CcValue::Null),
            CcValue::Str(_) => Ok(CcValue::Double(self.num()?.value)),
            CcValue::Boolean(_) | CcValue::Empty => {
                let n = self.num()?;
                CcValue::make(Rank::Integer, n.value)
            }
            _ => {
                self.num()?;
                Ok(self.clone())
            }
        }
    }

    /// `Not`: the bitwise complement of the operand rounded to a Long.
    pub fn not(&self) -> CcResult {
        match self {
            CcValue::Null => Ok(CcValue::Null),
            CcValue::Boolean(b) => Ok(CcValue::Boolean(!b)),
            _ => {
                let rank = self.num()?.rank;
                let v = !self.long()?;
                Ok(if rank == Rank::Integer { CcValue::Integer(v as i16) } else { CcValue::Long(v) })
            }
        }
    }

    /// Apply a binary operator to two operands that are written out in the directive.
    pub fn binary(op: CcOp, l: &CcValue, r: &CcValue) -> CcResult {
        Self::binary_with(op, l, r, false)
    }

    /// Apply a binary operator. `by_name` says an operand was reached through a constant's
    /// name: a number compared with text is then compared as text (its own text against the
    /// other), where operands written out as literals need the text to be a number.
    pub fn binary_with(op: CcOp, l: &CcValue, r: &CcValue, by_name: bool) -> CcResult {
        use CcOp::*;
        match op {
            Is => {
                return if *l == CcValue::Nothing && *r == CcValue::Nothing { Ok(CcValue::TRUE) } else { Err(ERR_TYPE_MISMATCH) };
            }
            Cat => {
                return Ok(CcValue::Str(format!("{}{}", l.text()?, r.text()?)));
            }
            And | Or => {
                if let Some(v) = null_logic(op, l, r)? {
                    return Ok(v);
                }
            }
            _ => {}
        }
        if *l == CcValue::Null || *r == CcValue::Null {
            // A comparison or operation on Null is Null (the operand rules above aside).
            return Ok(CcValue::Null);
        }
        match op {
            Add => {
                if let (CcValue::Str(a), CcValue::Str(b)) = (l, r) {
                    return Ok(CcValue::Str(format!("{a}{b}")));
                }
                arithmetic(l, r, |a, b| a + b, |a, b| a.checked_add(b))
            }
            Sub => arithmetic(l, r, |a, b| a - b, |a, b| a.checked_sub(b)),
            Mul => {
                let (a, b) = (l.num()?, r.num()?);
                if a.rank == Rank::Currency && b.rank == Rank::Currency {
                    return CcValue::make(Rank::Currency, a.value * b.value);
                }
                arithmetic(l, r, |a, b| a * b, |_, _| None)
            }
            Div => {
                let (a, b) = (l.num()?, r.num()?);
                if b.value == 0.0 {
                    return Err(ERR_DIVISION_BY_ZERO);
                }
                CcValue::make(Rank::Double, a.value / b.value)
            }
            IDiv | Mod => {
                let (a, b) = (l.long()?, r.long()?);
                if b == 0 {
                    return Err(ERR_DIVISION_BY_ZERO);
                }
                let v = if op == IDiv { a.checked_div(b) } else { a.checked_rem(b) }.ok_or(ERR_OVERFLOW)?;
                let small = l.num()?.rank == Rank::Integer && r.num()?.rank == Rank::Integer;
                Ok(if small { CcValue::Integer(v as i16) } else { CcValue::Long(v) })
            }
            Pow => {
                let (a, b) = (l.num()?.value, r.num()?.value);
                let v = a.powf(b);
                if (a == 0.0 && b < 0.0) || v.is_nan() {
                    return Err(ERR_INVALID_PROCEDURE_CALL);
                }
                Ok(CcValue::Double(v))
            }
            Eq | Ne | Lt | Gt | Le | Ge => {
                let one_text = matches!(l, CcValue::Str(_)) != matches!(r, CcValue::Str(_));
                let ord = if by_name && one_text {
                    l.text()?.to_lowercase().cmp(&r.text()?.to_lowercase())
                } else {
                    compare(l, r)?
                };
                Ok(CcValue::from_bool(match op {
                    Eq => ord == std::cmp::Ordering::Equal,
                    Ne => ord != std::cmp::Ordering::Equal,
                    Lt => ord == std::cmp::Ordering::Less,
                    Gt => ord == std::cmp::Ordering::Greater,
                    Le => ord != std::cmp::Ordering::Greater,
                    _ => ord != std::cmp::Ordering::Less,
                }))
            }
            Like => Ok(CcValue::from_bool(like(&l.text()?, &r.text()?))),
            And | Or | Xor | Eqv | Imp => {
                let (a, b) = (l.long()?, r.long()?);
                let v = match op {
                    And => a & b,
                    Or => a | b,
                    Xor => a ^ b,
                    Eqv => !(a ^ b),
                    _ => !a | b,
                };
                Ok(logic_result(l, r, v))
            }
            Cat | Is => unreachable!(),
        }
    }

    /// Whether `name` is one of the built-in functions a directive may call.
    pub fn is_function_name(name: &str) -> bool {
        FUNCTIONS.contains(&name.trim_end_matches('$').to_ascii_lowercase().as_str())
    }

    /// A call of one of the built-in functions a directive may use, or `None` for any
    /// other name.
    pub fn call(name: &str, args: &[CcValue]) -> Option<CcResult> {
        let lower = name.trim_end_matches('$').to_ascii_lowercase();
        if !FUNCTIONS.contains(&lower.as_str()) {
            return None;
        }
        if args.len() != 1 {
            return Some(Err(ERR_INVALID_PROCEDURE_CALL));
        }
        Some(call_one(&lower, &args[0]))
    }
}

/// `And` / `Or` with a Null operand: the other operand decides when it can.
fn null_logic(op: CcOp, l: &CcValue, r: &CcValue) -> Result<Option<CcValue>, u32> {
    let (null, other) = match (l, r) {
        (CcValue::Null, CcValue::Null) => return Ok(Some(CcValue::Null)),
        (CcValue::Null, o) | (o, CcValue::Null) => (true, o),
        _ => return Ok(None),
    };
    let _ = null;
    let v = other.long()?;
    Ok(Some(match (op, v) {
        (CcOp::And, 0) => logic_result(other, other, 0),
        (CcOp::Or, v) if v != 0 => logic_result(other, other, v),
        _ => CcValue::Null,
    }))
}

/// The type of a bitwise result: Boolean from two Booleans, Integer from operands no wider
/// than Integer, otherwise Long.
fn logic_result(l: &CcValue, r: &CcValue, v: i32) -> CcValue {
    if matches!((l, r), (CcValue::Boolean(_), CcValue::Boolean(_))) {
        return CcValue::Boolean(v != 0);
    }
    let small = |x: &CcValue| x.num().map_or(false, |n| n.rank == Rank::Integer);
    if small(l) && small(r) { CcValue::Integer(v as i16) } else { CcValue::Long(v) }
}

/// `+`, `-`, `*`: the result has the wider operand's type and must fit it.
fn arithmetic(l: &CcValue, r: &CcValue, f: fn(f64, f64) -> f64, exact: fn(i64, i64) -> Option<i64>) -> CcResult {
    let (a, b) = (l.num()?, r.num()?);
    // Currency with whole numbers stays Currency, and is exact.
    let (a, b) = match (a.rank, b.rank) {
        (Rank::Currency, Rank::Integer | Rank::Long | Rank::Currency) | (Rank::Integer | Rank::Long, Rank::Currency) => {
            let sa = a.scaled.unwrap_or((a.value * 10_000.0) as i64);
            let sb = b.scaled.unwrap_or((b.value * 10_000.0) as i64);
            if let Some(v) = exact(sa, sb) {
                return Ok(CcValue::Currency(v));
            }
            (a, b)
        }
        _ => (a, b),
    };
    let rank = if a.rank == Rank::Currency || b.rank == Rank::Currency {
        if a.rank >= Rank::Single || b.rank >= Rank::Single { a.rank.max_with(b.rank) } else { Rank::Currency }
    } else {
        a.rank.max_with(b.rank)
    };
    CcValue::make(rank, f(a.value, b.value))
}

impl Rank {
    fn max_with(self, other: Rank) -> Rank {
        if self >= other { self } else { other }
    }
}

fn boolean_value(b: bool) -> i8 {
    if b { -1 } else { 0 }
}

/// The Boolean a text stands for (as a number: -1 or 0).
fn text_as_boolean(text: &str) -> Result<i8, u32> {
    let word = text.trim();
    if word.eq_ignore_ascii_case("true") {
        return Ok(-1);
    }
    if word.eq_ignore_ascii_case("false") {
        return Ok(0);
    }
    parse_number(text).map(|v| boolean_value(v != 0.0)).ok_or(ERR_TYPE_MISMATCH)
}

/// Variant comparison: two strings compare as text (case-insensitively); a string against
/// a number needs the string to be numeric; Empty is 0 against a number and "" against
/// text.
fn compare(l: &CcValue, r: &CcValue) -> Result<std::cmp::Ordering, u32> {
    use std::cmp::Ordering;
    match (l, r) {
        (CcValue::Str(a), CcValue::Str(b)) => Ok(a.to_lowercase().cmp(&b.to_lowercase())),
        (CcValue::Empty, CcValue::Str(b)) => Ok(String::new().cmp(&b.to_lowercase())),
        (CcValue::Str(a), CcValue::Empty) => Ok(a.to_lowercase().cmp(&String::new())),
        // Text against a Boolean is read as a Boolean: `True` or `False` in any case, or a number.
        (CcValue::Boolean(b), CcValue::Str(t)) => Ok(boolean_value(*b).cmp(&text_as_boolean(t)?)),
        (CcValue::Str(t), CcValue::Boolean(b)) => Ok(text_as_boolean(t)?.cmp(&boolean_value(*b))),
        _ => {
            let (a, b) = (l.num()?, r.num()?);
            if let (Some(x), Some(y)) = (a.scaled, b.scaled) {
                return Ok(x.cmp(&y));
            }
            Ok(a.value.partial_cmp(&b.value).unwrap_or(Ordering::Equal))
        }
    }
}

/// The result of one of the built-in functions on one argument.
fn call_one(name: &str, v: &CcValue) -> CcResult {
    match name {
        "len" | "lenb" => {
            if *v == CcValue::Null {
                return Ok(CcValue::Null);
            }
            let n = v.text()?.chars().count() as i32;
            Ok(CcValue::Long(if name == "lenb" { n * 2 } else { n }))
        }
        "cvar" => Ok(v.clone()),
        "cstr" => Ok(CcValue::Str(v.text()?)),
        "cbool" => Ok(CcValue::Boolean(v.is_true()?)),
        "abs" | "int" | "fix" | "sgn" => {
            let n = v.num()?;
            let value = match name {
                "abs" => n.value.abs(),
                "int" => n.value.floor(),
                "fix" => n.value.trunc(),
                _ => return Ok(CcValue::Integer(if n.value > 0.0 { 1 } else if n.value < 0.0 { -1 } else { 0 })),
            };
            let rank = match (n.rank, matches!(v, CcValue::Str(_))) {
                (_, true) => Rank::Double,
                (Rank::Integer, _) if value > i16::MAX as f64 => Rank::Long,
                (rank, _) => rank,
            };
            if rank == Rank::Currency {
                return CcValue::make(Rank::Currency, value);
            }
            CcValue::make(rank, value)
        }
        "cbyte" => {
            let n = round_half_even(v.num()?.value);
            if (0.0..=255.0).contains(&n) { Ok(CcValue::Integer(n as i16)) } else { Err(ERR_OVERFLOW) }
        }
        "cint" => CcValue::make(Rank::Integer, round_half_even(v.num()?.value)),
        "clng" => CcValue::make(Rank::Long, round_half_even(v.num()?.value)),
        "csng" => CcValue::make(Rank::Single, v.num()?.value),
        "cdbl" => Ok(CcValue::Double(v.num()?.value)),
        "ccur" => CcValue::make(Rank::Currency, v.num()?.value),
        "cdate" => Ok(CcValue::Date(v.num()?.value)),
        // "cverr"
        _ => {
            let n = round_half_even(v.num()?.value);
            if (0.0..=65535.0).contains(&n) { Ok(CcValue::Error(n as i32)) } else { Err(ERR_INVALID_PROCEDURE_CALL) }
        }
    }
}

/// `Like` pattern matching (`?`, `*`, `#`, `[list]`, `[!list]`), ignoring case.
fn like(text: &str, pattern: &str) -> bool {
    let t: Vec<char> = text.to_lowercase().chars().collect();
    let p: Vec<char> = pattern.to_lowercase().chars().collect();
    like_from(&t, &p)
}

fn like_from(t: &[char], p: &[char]) -> bool {
    let Some(&first) = p.first() else { return t.is_empty() };
    match first {
        '*' => (0..=t.len()).any(|i| like_from(&t[i..], &p[1..])),
        '?' => !t.is_empty() && like_from(&t[1..], &p[1..]),
        '#' => t.first().is_some_and(|c| c.is_ascii_digit()) && like_from(&t[1..], &p[1..]),
        '[' => {
            let Some(close) = p.iter().position(|&c| c == ']') else {
                return t.first() == Some(&'[') && like_from(&t[1..], &p[1..]);
            };
            let set = &p[1..close];
            let (negate, set) = match set.first() {
                Some('!') => (true, &set[1..]),
                _ => (false, set),
            };
            let Some(&c) = t.first() else { return false };
            let mut hit = false;
            let mut i = 0;
            while i < set.len() {
                if i + 2 < set.len() && set[i + 1] == '-' {
                    hit |= set[i] <= c && c <= set[i + 2];
                    i += 3;
                } else {
                    hit |= set[i] == c;
                    i += 1;
                }
            }
            hit != negate && like_from(&t[1..], &p[close + 1..])
        }
        c => t.first() == Some(&c) && like_from(&t[1..], &p[1..]),
    }
}

/// A number written as text: decimal (with exponent) or `&H` / `&O`, with blanks around.
fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    let lower = t.to_ascii_lowercase();
    if let Some(h) = lower.strip_prefix("&h") {
        return i64::from_str_radix(h.trim_end_matches('&'), 16).ok().map(|v| radix_value(v, h.len(), 4));
    }
    if let Some(o) = lower.strip_prefix("&o") {
        return i64::from_str_radix(o.trim_end_matches('&'), 8).ok().map(|v| radix_value(v, o.len(), 6));
    }
    let ok = lower.chars().all(|c| c.is_ascii_digit() || matches!(c, '.' | 'e' | 'd' | '+' | '-'));
    if !ok || !lower.chars().any(|c| c.is_ascii_digit()) {
        return None;
    }
    lower.replace('d', "e").parse::<f64>().ok()
}

/// A hexadecimal or octal number of up to four (six) digits is an Integer, so one that
/// sets the top bit is negative.
fn radix_value(v: i64, digits: usize, integer_digits: usize) -> f64 {
    if digits <= integer_digits && v >= 0x8000 && v <= 0xFFFF { (v - 0x10000) as f64 } else { v as f64 }
}

fn round_half_even(n: f64) -> f64 {
    let r = n.round();
    if (n - n.trunc()).abs() == 0.5 && r % 2.0 != 0.0 {
        r - n.signum()
    } else {
        r
    }
}

/// A Double (Single) as `&` shows it: at most `sig` significant digits, scientific
/// below 1E-4 or from 1E`sig`.
fn format_float(d: f64, sig: usize) -> String {
    if d == 0.0 {
        return "0".into();
    }
    if !d.is_finite() {
        return if d.is_nan() { "NaN".into() } else if d > 0.0 { "1.#INF".into() } else { "-1.#INF".into() };
    }
    let sci = format!("{:.*e}", sig - 1, d);
    let (mant, exp) = sci.split_once('e').unwrap();
    let exp: i32 = exp.parse().unwrap();
    let neg = mant.starts_with('-');
    let digits: String = mant.chars().filter(|c| c.is_ascii_digit()).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let sign = if neg { "-" } else { "" };
    if exp < -4 || exp >= sig as i32 {
        let m = if digits.len() > 1 { format!("{}.{}", &digits[..1], &digits[1..]) } else { digits.to_string() };
        return format!("{sign}{m}E{}{:02}", if exp < 0 { '-' } else { '+' }, exp.abs());
    }
    if exp >= 0 {
        let int_len = exp as usize + 1;
        if digits.len() <= int_len {
            format!("{sign}{digits}{}", "0".repeat(int_len - digits.len()))
        } else {
            format!("{sign}{}.{}", &digits[..int_len], &digits[int_len..])
        }
    } else {
        format!("{sign}0.{}{digits}", "0".repeat((-exp - 1) as usize))
    }
}

fn format_currency(c: i64) -> String {
    let whole = c / 10_000;
    let frac = (c % 10_000).abs();
    let sign = if c < 0 && whole == 0 { "-" } else { "" };
    if frac == 0 {
        format!("{sign}{whole}")
    } else {
        format!("{sign}{whole}.{}", format!("{frac:04}").trim_end_matches('0'))
    }
}

/// Why the project's conditional-compilation arguments are refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConstantDeclarationsError {
    /// The text is not written as the project format requires ("Invalid syntax for constant
    /// declarations").
    InvalidSyntax,
    /// The text is 1024 characters or longer ("Fixed or static data can't be larger than 64K").
    DataTooLarge,
}

/// The longest argument text, in 16-bit units, that is read at all.
const MAX_ARGUMENT_UNITS: usize = 1023;

/// A blank in the argument text: the tab, line end, vertical tab and form feed characters, and
/// the space. No other character counts, the no-break space included.
fn is_argument_blank(c: char) -> bool {
    matches!(c, '\t'..='\r' | ' ')
}

/// A letter in a name: an ASCII letter, or any character from U+0080 to U+00FF.
fn is_name_letter(c: char) -> bool {
    c.is_ascii_alphabetic() || ('\u{80}'..='\u{ff}').contains(&c)
}

/// A constant's name: a letter, then letters, digits and underscores. An underscore cannot
/// start it.
fn is_constant_name(name: &[char]) -> bool {
    match name.split_first() {
        Some((first, rest)) => {
            is_name_letter(*first) && rest.iter().all(|c| is_name_letter(*c) || c.is_ascii_digit() || *c == '_')
        }
        None => false,
    }
}

/// The form of a name in which two names that compare equal ignoring case are identical: the
/// letters fold to lower case, and the sharp s and the ligatures AE and TH compare equal to
/// the letter pairs they stand for.
fn name_comparison_key(name: &str) -> String {
    let mut key = String::new();
    for c in crate::frontend::scanner::fold_name(name).chars() {
        match c {
            '\u{df}' => key.push_str("ss"),
            '\u{e6}' => key.push_str("ae"),
            '\u{fe}' => key.push_str("th"),
            _ => key.push(c),
        }
    }
    key
}

/// Split the project's conditional-compilation arguments (`DEBUG = 1 : Level = -2`) into
/// `(folded name, value)` pairs.
///
/// Each argument is a name (see [`is_constant_name`]; a reserved word will do), an `=`, and a
/// 16-bit signed integer. Blanks are skipped before and after each part, and between a sign and
/// its digits. The value is an optional `+` or `-` and then digits; a sign alone is the value 0.
/// A run of digits is read as a 32-bit number, wrapping, and must then fit 16 bits. Without a
/// sign, five digits are the most a value may have, unless an earlier value (or this one) had a
/// sign, after which a run may be up to eleven digits long, signed or not. A name given twice
/// (ignoring case), an empty part, or any other text makes the whole text invalid; a text of
/// 1024 or more characters is refused before it is read.
pub fn parse_cond_comp_args(text: &str) -> Result<Vec<(String, CcValue)>, ConstantDeclarationsError> {
    use ConstantDeclarationsError::{DataTooLarge, InvalidSyntax};
    if text.encode_utf16().count() > MAX_ARGUMENT_UNITS {
        return Err(DataTooLarge);
    }
    let chars: Vec<char> = text.chars().collect();
    let at = |i: usize| chars.get(i).copied();
    let skip_blanks = |mut i: usize| {
        while at(i).is_some_and(is_argument_blank) {
            i += 1;
        }
        i
    };
    let mut out: Vec<(String, CcValue)> = Vec::new();
    let mut pos = skip_blanks(0);
    if pos == chars.len() {
        return Ok(out);
    }
    let entries = 1 + chars[pos..].iter().filter(|c| **c == ':').count();
    let mut sign_seen = false;
    for entry in 1..=entries {
        // The name runs to the next `=` (which may lie beyond a `:`, and then is no name).
        let eq = (pos..chars.len()).find(|i| chars[*i] == '=').filter(|eq| *eq > pos).ok_or(InvalidSyntax)?;
        let mut name_end = eq;
        while name_end > pos && is_argument_blank(chars[name_end - 1]) {
            name_end -= 1;
        }
        let name = &chars[pos..name_end];
        if !is_constant_name(name) {
            return Err(InvalidSyntax);
        }
        let name: String = name.iter().collect();
        pos = skip_blanks(eq + 1);
        if pos == chars.len() {
            return Err(InvalidSyntax);
        }
        let negative = at(pos) == Some('-');
        let signed = negative || at(pos) == Some('+');
        if signed {
            sign_seen = true;
            let after_sign = skip_blanks(pos + 1);
            // A sign alone is a value; blanks after it must lead to something.
            if after_sign > pos + 1 && after_sign == chars.len() {
                return Err(InvalidSyntax);
            }
            pos = after_sign;
        }
        let digits_start = pos;
        while at(pos).is_some_and(|c| c.is_ascii_digit()) {
            pos += 1;
        }
        let digit_count = pos - digits_start;
        if digit_count == 0 && !signed {
            return Err(InvalidSyntax);
        }
        let longest_run = if sign_seen { 11 } else { 5 };
        if digit_count > longest_run {
            return Err(InvalidSyntax);
        }
        let magnitude = chars[digits_start..pos]
            .iter()
            .fold(0u32, |acc, c| acc.wrapping_mul(10).wrapping_add(c.to_digit(10).unwrap_or(0)));
        let number = if negative { (magnitude as i32).wrapping_neg() } else { magnitude as i32 };
        let number = i16::try_from(number).map_err(|_| InvalidSyntax)?;
        let key = name_comparison_key(&name);
        if out.iter().any(|(n, _)| name_comparison_key(n) == key) {
            return Err(InvalidSyntax);
        }
        out.push((crate::frontend::scanner::fold_name(&name), CcValue::Integer(number)));
        pos = skip_blanks(pos);
        if entry == entries {
            if pos != chars.len() {
                return Err(InvalidSyntax);
            }
        } else {
            if at(pos) != Some(':') {
                return Err(InvalidSyntax);
            }
            pos = skip_blanks(pos + 1);
            if pos == chars.len() {
                return Err(InvalidSyntax);
            }
        }
    }
    Ok(out)
}
