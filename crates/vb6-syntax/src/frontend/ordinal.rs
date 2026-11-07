//! The text after the `#` of a `Declare ... Alias "#..."` names an ordinal. It is read in two
//! steps: the number scanner must find a whole number at its start (see
//! [`crate::frontend::scanner::Scanner::ordinal_start`]), and then the whole text must convert
//! to an Integer the way a string converts to an Integer in the language runtime: blanks around
//! it, a sign before or after it, thousands separators, `&H`/`&O` forms, a trailing decimal
//! point, an exponent, within the Integer range.

/// Whether `text` (the alias text after the `#`) converts to an Integer.
pub fn converts_to_integer(text: &str) -> bool {
    let chars: Vec<char> = text.chars().collect();
    parse(&chars).is_some()
}

fn is_blank(c: char) -> bool {
    matches!(c, ' ' | '\t' | '\n' | '\r' | '\u{a0}')
}

/// A read position over the alias text.
struct Cursor<'a> {
    s: &'a [char],
    i: usize,
}

impl Cursor<'_> {
    fn peek(&self) -> Option<char> {
        self.s.get(self.i).copied()
    }

    fn skip_blanks(&mut self) {
        while self.peek().is_some_and(is_blank) {
            self.i += 1;
        }
    }

    fn at_end(&self) -> bool {
        self.i == self.s.len()
    }

    /// Consumes a `+` or `-` and says whether it was a `-`.
    fn take_sign(&mut self) -> Option<bool> {
        let negative = match self.peek() {
            Some('+') => false,
            Some('-') => true,
            _ => return None,
        };
        self.i += 1;
        Some(negative)
    }

    /// Consumes the digits of `radix` and returns their value (saturating), or `None` when
    /// there are no digits.
    fn take_radix_digits(&mut self, radix: u32) -> Option<u64> {
        let start = self.i;
        let mut value: u64 = 0;
        while let Some(d) = self.peek().and_then(|c| c.to_digit(radix)) {
            value = value.saturating_mul(radix as u64).saturating_add(d as u64);
            self.i += 1;
        }
        (self.i != start).then_some(value)
    }

    /// Digits, with thousands separators among them.
    fn take_integer_digits(&mut self) -> String {
        let mut digits = String::new();
        while let Some(c) = self.peek() {
            match c {
                '0'..='9' => digits.push(c),
                ',' if !digits.is_empty() => {}
                _ => break,
            }
            self.i += 1;
        }
        digits
    }

    /// A decimal point and the digits after it.
    fn take_fraction(&mut self) -> String {
        let mut fraction = String::new();
        if self.peek() != Some('.') {
            return fraction;
        }
        self.i += 1;
        while let Some(c @ '0'..='9') = self.peek() {
            fraction.push(c);
            self.i += 1;
        }
        fraction
    }

    /// An exponent (0 when absent); `None` when its marker is not followed by digits.
    fn take_exponent(&mut self) -> Option<i32> {
        if !matches!(self.peek(), Some('e' | 'E' | 'd' | 'D')) {
            return Some(0);
        }
        self.i += 1;
        let exponent_negative = self.take_sign().unwrap_or(false);
        let start = self.i;
        let mut exponent: i32 = 0;
        while let Some(d) = self.peek().and_then(|c| c.to_digit(10)) {
            exponent = exponent.saturating_mul(10).saturating_add(d as i32).min(10_000);
            self.i += 1;
        }
        if self.i == start {
            return None;
        }
        Some(if exponent_negative { -exponent } else { exponent })
    }
}

/// Hexadecimal and octal forms: the value keeps its sixteen-bit pattern. The cursor is just
/// after the `&`.
fn parse_radix_form(cur: &mut Cursor) -> Option<i64> {
    let radix = match cur.peek() {
        Some('h' | 'H') => {
            cur.i += 1;
            16
        }
        Some('o' | 'O') => {
            cur.i += 1;
            8
        }
        _ => 8,
    };
    let value = cur.take_radix_digits(radix)?;
    if value > 0xFFFF {
        return None;
    }
    cur.skip_blanks();
    cur.at_end().then_some(value as u16 as i16 as i64)
}

/// The value `digits.fraction * 10^exponent` converts to the nearest whole number, a half going
/// to the even one, within the Integer range.
fn round_to_integer(digits: &str, fraction: &str, exponent: i32, negative: bool) -> Option<i64> {
    let text = format!("{digits}.{}e{exponent}", if fraction.is_empty() { "0" } else { fraction });
    let magnitude: f64 = text.parse().ok()?;
    let value = if negative { -magnitude } else { magnitude };
    let mut rounded = value.round();
    if (value - value.trunc()).abs() == 0.5 && rounded % 2.0 != 0.0 {
        rounded -= value.signum();
    }
    (-32768.0..=32767.0).contains(&rounded).then_some(rounded as i64)
}

/// Value of the text as a whole number in the Integer range, or `None` when the text is not
/// a number or does not fit.
fn parse(s: &[char]) -> Option<i64> {
    let mut cur = Cursor { s, i: 0 };
    cur.skip_blanks();
    if cur.peek() == Some('&') {
        cur.i += 1;
        return parse_radix_form(&mut cur);
    }
    let leading_sign = cur.take_sign();
    let mut negative = leading_sign.unwrap_or(false);
    let digits = cur.take_integer_digits();
    if digits.is_empty() {
        return None;
    }
    let fraction = cur.take_fraction();
    let exponent = cur.take_exponent()?;
    cur.skip_blanks();
    // A sign may follow the number instead of leading it.
    if leading_sign.is_none() {
        if let Some(trailing) = cur.take_sign() {
            negative = trailing;
            cur.skip_blanks();
        }
    }
    if !cur.at_end() {
        return None;
    }
    round_to_integer(&digits, &fraction, exponent, negative)
}
