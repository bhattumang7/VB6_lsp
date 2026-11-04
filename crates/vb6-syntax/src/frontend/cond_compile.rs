//! Values of conditional-compilation expressions (`#Const`, `#If`) and the
//! project's conditional-compilation arguments.
//!
//! A conditional-compilation value behaves as a Variant restricted to Empty,
//! numbers and strings: operators coerce as VB does, comparisons yield
//! `True` (-1) or `False` (0), and `And`/`Or`/`Not`/... act on the bits of
//! the operands rounded to Long.

/// A conditional-compilation value.
#[derive(Debug, Clone, PartialEq)]
pub enum CcValue {
    Empty,
    /// An Integer or Long (and `True` = -1, `False` = 0).
    Int(i32),
    Double(f64),
    Str(String),
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
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

impl CcValue {
    pub const TRUE: CcValue = CcValue::Int(-1);
    pub const FALSE: CcValue = CcValue::Int(0);

    fn from_bool(b: bool) -> CcValue {
        if b { CcValue::TRUE } else { CcValue::FALSE }
    }

    /// The numeric value, if the value is a number, Empty, or numeric text.
    pub fn number(&self) -> Option<f64> {
        match self {
            CcValue::Empty => Some(0.0),
            CcValue::Int(n) => Some(*n as f64),
            CcValue::Double(d) => Some(*d),
            CcValue::Str(s) => parse_number(s),
        }
    }

    /// The text of the value, as `&` and string comparison see it.
    pub fn text(&self) -> String {
        match self {
            CcValue::Empty => String::new(),
            CcValue::Int(n) => n.to_string(),
            CcValue::Double(d) => {
                if d.fract() == 0.0 && d.abs() < 1e15 { format!("{}", *d as i64) } else { format!("{d}") }
            }
            CcValue::Str(s) => s.clone(),
        }
    }

    /// The value rounded (halves to even) to a Long, for the bitwise
    /// operators, `\` and `Mod`.
    pub fn long(&self) -> i32 {
        let n = self.number().unwrap_or(0.0);
        round_half_even(n).clamp(i32::MIN as f64, i32::MAX as f64) as i32
    }

    /// Whether a condition with this value is taken: a non-zero number.
    pub fn is_true(&self) -> bool {
        self.number().is_some_and(|n| n != 0.0)
    }

    /// A numeric result: an Int when it is a whole number in Long range,
    /// otherwise a Double.
    fn from_f64(n: f64) -> CcValue {
        if n.fract() == 0.0 && n >= i32::MIN as f64 && n <= i32::MAX as f64 {
            CcValue::Int(n as i32)
        } else {
            CcValue::Double(n)
        }
    }

    /// Unary minus.
    pub fn negate(&self) -> CcValue {
        CcValue::from_f64(-self.number().unwrap_or(0.0))
    }

    /// `Not`: the bitwise complement of the operand rounded to a Long.
    pub fn not(&self) -> CcValue {
        CcValue::Int(!self.long())
    }

    /// Apply a binary operator.
    pub fn binary(op: CcOp, l: &CcValue, r: &CcValue) -> CcValue {
        use CcOp::*;
        let num = |v: &CcValue| v.number().unwrap_or(0.0);
        match op {
            Add => match (l, r) {
                (CcValue::Str(a), CcValue::Str(b)) => CcValue::Str(format!("{a}{b}")),
                _ => CcValue::from_f64(num(l) + num(r)),
            },
            Sub => CcValue::from_f64(num(l) - num(r)),
            Mul => CcValue::from_f64(num(l) * num(r)),
            // The quotient of `/` is a floating-point number; a zero divisor
            // gives 0 here, as there is no error channel in a directive.
            Div => {
                let d = num(r);
                if d == 0.0 { CcValue::Int(0) } else { CcValue::from_f64(num(l) / d) }
            }
            IDiv => {
                let d = r.long();
                if d == 0 { CcValue::Int(0) } else { CcValue::Int(l.long().wrapping_div(d)) }
            }
            Mod => {
                let d = r.long();
                if d == 0 { CcValue::Int(0) } else { CcValue::Int(l.long().wrapping_rem(d)) }
            }
            Pow => CcValue::from_f64(num(l).powf(num(r))),
            Cat => CcValue::Str(format!("{}{}", l.text(), r.text())),
            Eq | Ne | Lt | Gt | Le | Ge => {
                let ord = compare(l, r);
                CcValue::from_bool(match op {
                    Eq => ord == std::cmp::Ordering::Equal,
                    Ne => ord != std::cmp::Ordering::Equal,
                    Lt => ord == std::cmp::Ordering::Less,
                    Gt => ord == std::cmp::Ordering::Greater,
                    Le => ord != std::cmp::Ordering::Greater,
                    _ => ord != std::cmp::Ordering::Less,
                })
            }
            And => CcValue::Int(l.long() & r.long()),
            Or => CcValue::Int(l.long() | r.long()),
            Xor => CcValue::Int(l.long() ^ r.long()),
            Eqv => CcValue::Int(!(l.long() ^ r.long())),
            Imp => CcValue::Int(!l.long() | r.long()),
        }
    }
}

/// Variant comparison: two strings compare as text (binary); a string
/// against a number compares as numbers when the string is numeric, and is
/// otherwise the greater; Empty is 0 against a number and "" against text.
fn compare(l: &CcValue, r: &CcValue) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    match (l, r) {
        (CcValue::Str(a), CcValue::Str(b)) => a.cmp(b),
        (CcValue::Empty, CcValue::Str(b)) => String::new().cmp(b),
        (CcValue::Str(a), CcValue::Empty) => a.cmp(&String::new()),
        (CcValue::Str(a), other) => match (parse_number(a), other.number()) {
            (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(Ordering::Equal),
            _ => Ordering::Greater,
        },
        (other, CcValue::Str(b)) => match (other.number(), parse_number(b)) {
            (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(Ordering::Equal),
            _ => Ordering::Less,
        },
        _ => l.number().unwrap_or(0.0).partial_cmp(&r.number().unwrap_or(0.0)).unwrap_or(Ordering::Equal),
    }
}

fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

fn round_half_even(n: f64) -> f64 {
    let r = n.round();
    if (n - n.trunc()).abs() == 0.5 && r % 2.0 != 0.0 {
        r - n.signum()
    } else {
        r
    }
}

/// Split the project's conditional-compilation arguments
/// (`DEBUG = 1 : Name = "text"`) into `(lower-cased name, value)` pairs. A
/// value is an integer, a number, or a double-quoted string (`""` for a
/// quote); a missing or unreadable value is Empty. Colons inside a quoted
/// string do not split.
pub fn parse_cond_comp_args(text: &str) -> Vec<(String, CcValue)> {
    let mut out = Vec::new();
    let mut parts: Vec<String> = Vec::new();
    let mut cur = String::new();
    let mut in_str = false;
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' if in_str && chars.peek() == Some(&'"') => {
                cur.push_str("\"\"");
                chars.next();
            }
            '"' => {
                in_str = !in_str;
                cur.push(c);
            }
            ':' if !in_str => parts.push(std::mem::take(&mut cur)),
            _ => cur.push(c),
        }
    }
    parts.push(cur);
    for part in parts {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }
        let (name, value) = match part.split_once('=') {
            Some((n, v)) => (n.trim(), v.trim()),
            None => (part, ""),
        };
        if name.is_empty() {
            continue;
        }
        let value = if let Some(body) = value.strip_prefix('"').and_then(|v| v.strip_suffix('"')) {
            CcValue::Str(body.replace("\"\"", "\""))
        } else if value.is_empty() {
            CcValue::Empty
        } else if let Ok(n) = value.parse::<i32>() {
            CcValue::Int(n)
        } else if let Ok(d) = value.parse::<f64>() {
            CcValue::Double(d)
        } else if value.eq_ignore_ascii_case("true") {
            CcValue::TRUE
        } else if value.eq_ignore_ascii_case("false") {
            CcValue::FALSE
        } else {
            CcValue::Empty
        };
        out.push((name.to_ascii_lowercase(), value));
    }
    out
}
