//! Typed values in directive expressions: arithmetic that overflows its type, the errors
//! an expression raises (reported on the directive), Null, the built-in functions, `Like`
//! and `Is`, text formatting, and operator precedence.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::ExprNode;

const OVERFLOW: u32 = 6;
const INVALID_PROCEDURE_CALL: u32 = 5;
const DIVISION_BY_ZERO: u32 = 11;
const TYPE_MISMATCH: u32 = 13;
const INVALID_USE_OF_NULL: u32 = 94;
const VARIABLE_NOT_DEFINED: u32 = 0x9caf;

fn run(cond: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("#If {cond} Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n");
    let p = parse(&src);
    let name = match p.top.first().map(|n| p.node(*n)) {
        Some(ExprNode::ProcDecl { name, .. }) => p.name(*name).to_string(),
        _ => String::new(),
    };
    (name, p.found)
}

fn yes(cond: &str) {
    assert_eq!(run(cond), ("Yes".to_string(), vec![]), "{cond}");
}
fn no(cond: &str) {
    assert_eq!(run(cond), ("No".to_string(), vec![]), "{cond}");
}
/// The directive is reported with `code` at its `#If`, and its branch is not taken.
fn fails(cond: &str, code: u32) {
    assert_eq!(run(cond), ("No".to_string(), vec![(code, 0, 3)]), "{cond}");
}

#[test]
fn arithmetic_overflows_the_type_of_its_operands() {
    for cond in [
        "32767 + 1 = 32768",
        "200 * 200",
        "2147483647 + 1",
        "100000 * 100000",
        "&H7FFFFFFF + 1",
        "CInt(40000)",
        "CLng(1E10)",
        "CByte(256)",
        "(1E10 Or 1)",
        "(1E10 \\ 2)",
    ] {
        fails(cond, OVERFLOW);
    }
    for cond in ["32768 + 1", "200& * 200", "32767 + 1&", "-32768 - 1 < 0", "1E308 * 10 > 0"] {
        yes(cond);
    }
}

#[test]
fn division_and_power_domain_errors() {
    for cond in ["1 / 0", "5 Mod 0", "7 \\ 0.4", "5 \\ 0"] {
        fails(cond, DIVISION_BY_ZERO);
    }
    for cond in ["0 ^ -1", "(-8) ^ (1 / 3)"] {
        fails(cond, INVALID_PROCEDURE_CALL);
    }
    let src = "#Const A = 1 / 0\nSub S()\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(DIVISION_BY_ZERO, 0, 6)]);
}

#[test]
fn text_that_is_not_a_number_is_a_type_mismatch() {
    for cond in ["\"a\"", "\"\"", "-\"a\"", "+\"a\"", "Not \"a\"", "\"a\" + 1", "\"a\" - 1", "1 < \"a\"", "\"a\" < 1", "CInt(\"a\")", "Abs(\"a\")"] {
        fails(cond, TYPE_MISMATCH);
    }
    for cond in ["\"1\"", "\" 1\"", "\"1 \"", "\".5\"", "\"1E2\" = 100", "\"&H10\" = 16", "\"1\" = 1", "\"inf\" = \"INF\""] {
        yes(cond);
    }
    fails("\"inf\" = 1", TYPE_MISMATCH);
}

#[test]
fn null_is_a_value_of_its_own() {
    for cond in ["Null", "Null = 1", "Null + 1", "Int(Null)", "Null Or 0"] {
        fails(cond, INVALID_USE_OF_NULL);
    }
    fails("Null & \"a\"", TYPE_MISMATCH);
    yes("Null Or 1");
    yes("(Null And 0) = 0");
    let src = "#Const A = Null\nSub S()\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![]);
}

#[test]
fn built_in_functions_and_other_calls() {
    for cond in [
        "Len(\"ab\") = 2",
        "LenB(\"ab\") = 4",
        "Abs(-3) = 3",
        "Abs(-32768) = 32768",
        "Int(-1.5) = -2",
        "Fix(-1.5) = -1",
        "Sgn(-3) = -1",
        "CBool(2)",
        "CByte(255) = 255",
        "CInt(2.5) = 2",
        "CLng(3.5) = 4",
        "CStr(1.5) = \"1.5\"",
        "CDbl(\"2\") = 2",
        "CSng(1.5) = 1.5",
        "CCur(0.1) + CCur(0.2) = CCur(0.3)",
        "CDate(1) = 1",
        "CVar(1) = 1",
    ] {
        yes(cond);
    }
    fails("CVErr(1)", TYPE_MISMATCH);
    for cond in ["Sqr(4)", "Chr(65)", "Asc(\"a\")", "Mid(\"abc\", 1)", "Foo(1)", "IIf(1, 2, 3)"] {
        fails(cond, VARIABLE_NOT_DEFINED);
    }
    for cond in ["Rnd", "Timer", "Now", "vbCrLf"] {
        no(cond);
    }
}

#[test]
fn like_and_is() {
    for cond in ["\"abc\" Like \"a*\"", "\"ABC\" Like \"a*\"", "\"a1\" Like \"a#\"", "\"b\" Like \"[a-c]\"", "\"d\" Like \"[!a-c]\"", "\"abc\" Like \"a?c\"", "Nothing Is Nothing"] {
        yes(cond);
    }
    no("\"abc\" Like \"b*\"");
    fails("1 Is 1", TYPE_MISMATCH);
}

#[test]
fn reserved_words_are_not_names_but_some_words_are() {
    for cond in ["Dim", "Print", "New", "String", "Integer", "Debug", "Resume", "End", "Tab"] {
        let (_, found) = run(cond);
        assert!(!found.is_empty(), "{cond}");
    }
    for cond in ["Me", "Error", "Name", "Line", "Date", "Time", "Text", "Step", "Width", "Mid"] {
        no(cond);
    }
}

#[test]
fn text_of_values() {
    for cond in [
        "True & \"x\" = \"Truex\"",
        "False & \"\" = \"False\"",
        "(1 = 1) & \"\" = \"True\"",
        "1E15 & \"\" = \"1E+15\"",
        "1E14 & \"\" = \"100000000000000\"",
        "(0.1 + 0.2) & \"\" = \"0.3\"",
        "0.00001 & \"\" = \"1E-05\"",
        "0.0001 & \"\" = \"0.0001\"",
        "-2.5 & \"\" = \"-2.5\"",
        "1.5@ & \"\" = \"1.5\"",
        "7 / 2 & \"\" = \"3.5\"",
    ] {
        yes(cond);
    }
}

#[test]
fn currency_is_exact() {
    yes("0.1@ + 0.2@ = 0.3@");
    no("0.1 + 0.2 = 0.3");
}

#[test]
fn predefined_names() {
    yes("Win32");
    yes("Win32 = -1");
    for name in ["VBA6", "VBA7", "Mac", "Win16", "Win64"] {
        no(name);
        yes(&format!("Not {name}"));
    }
}

#[test]
fn or_binds_tighter_than_xor_and_eqv_tighter_than_imp() {
    yes("(3 Xor 1 Or 1) = 2");
    yes("(0 Imp 0 Eqv 0) = -1");
    yes("(1 Or 2 And 0) = 1");
    yes("(Not 1 + 1) = -3");
    yes("-2 ^ 2 = -4");
    yes("2 * 3 Mod 4 = 2");
}
