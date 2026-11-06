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

/// Value of the text as a whole number in the Integer range, or `None` when the text is not
/// a number or does not fit.
fn parse(s: &[char]) -> Option<i64> {
    let mut i = 0;
    let skip_blanks = |i: &mut usize| {
        while *i < s.len() && is_blank(s[*i]) {
            *i += 1;
        }
    };
    skip_blanks(&mut i);
    // Hexadecimal and octal forms: the value keeps its sixteen-bit pattern.
    if s.get(i) == Some(&'&') {
        i += 1;
        let radix = match s.get(i) {
            Some('h' | 'H') => {
                i += 1;
                16
            }
            Some('o' | 'O') => {
                i += 1;
                8
            }
            _ => 8,
        };
        let start = i;
        let mut value: u64 = 0;
        while let Some(d) = s.get(i).and_then(|c| c.to_digit(radix)) {
            value = value.saturating_mul(radix as u64).saturating_add(d as u64);
            i += 1;
        }
        if i == start || value > 0xFFFF {
            return None;
        }
        skip_blanks(&mut i);
        return (i == s.len()).then_some(value as u16 as i16 as i64);
    }
    let mut negative = false;
    let mut signed = false;
    if let Some(sign @ ('+' | '-')) = s.get(i) {
        signed = true;
        negative = *sign == '-';
        i += 1;
    }
    // Digits, with thousands separators among them.
    let mut digits = String::new();
    while let Some(&c) = s.get(i) {
        match c {
            '0'..='9' => digits.push(c),
            ',' if !digits.is_empty() => {}
            _ => break,
        }
        i += 1;
    }
    if digits.is_empty() {
        return None;
    }
    // A decimal point and the digits after it.
    let mut fraction = String::new();
    if s.get(i) == Some(&'.') {
        i += 1;
        while let Some(&c @ '0'..='9') = s.get(i) {
            fraction.push(c);
            i += 1;
        }
    }
    // An exponent.
    let mut exponent: i32 = 0;
    if matches!(s.get(i), Some('e' | 'E' | 'd' | 'D')) {
        i += 1;
        let mut exponent_negative = false;
        if let Some(sign @ ('+' | '-')) = s.get(i) {
            exponent_negative = *sign == '-';
            i += 1;
        }
        let start = i;
        while let Some(d) = s.get(i).and_then(|c| c.to_digit(10)) {
            exponent = exponent.saturating_mul(10).saturating_add(d as i32).min(10_000);
            i += 1;
        }
        if i == start {
            return None;
        }
        if exponent_negative {
            exponent = -exponent;
        }
    }
    skip_blanks(&mut i);
    // A sign may follow the number instead of leading it.
    if !signed {
        if let Some(sign @ ('+' | '-')) = s.get(i) {
            negative = *sign == '-';
            i += 1;
            skip_blanks(&mut i);
        }
    }
    if i != s.len() {
        return None;
    }
    // The value is `digits.fraction * 10^exponent`; it converts to the nearest whole number,
    // a half going to the even one.
    let text = format!("{digits}.{}e{exponent}", if fraction.is_empty() { "0" } else { &fraction });
    let magnitude: f64 = text.parse().ok()?;
    let value = if negative { -magnitude } else { magnitude };
    let mut rounded = value.round();
    if (value - value.trunc()).abs() == 0.5 && rounded % 2.0 != 0.0 {
        rounded -= value.signum();
    }
    (-32768.0..=32767.0).contains(&rounded).then_some(rounded as i64)
}
