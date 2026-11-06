//! The names a module can use in a directive without defining them: the platform
//! constants, and what a name with no definition is.

mod common;
use common::parse;
use vb6_syntax::frontend::ast::{ExprArena, ExprNode};
use vb6_syntax::frontend::cond_compile::CcValue;
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

fn run(preamble: &str, cond: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("{preamble}#If {cond} Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n");
    let p = parse(&src);
    let name = match p.top.first().map(|n| p.node(*n)) {
        Some(ExprNode::ProcDecl { name, .. }) => p.name(*name).to_string(),
        _ => String::new(),
    };
    (name, p.found)
}

fn yes(cond: &str) {
    assert_eq!(run("", cond), ("Yes".to_string(), vec![]), "{cond} should be true");
}
fn no(cond: &str) {
    assert_eq!(run("", cond), ("No".to_string(), vec![]), "{cond} should be false");
}

#[test]
fn win32_is_the_integer_one() {
    yes("Win32");
    yes("Win32 = 1");
    no("Win32 = -1");
    no("Win32 = True");
    yes("Win32 + 1 = 2");
    yes("Win32 & \"a\" = \"1a\"");
}

#[test]
fn win16_and_mac_are_the_integer_zero() {
    for name in ["Win16", "Mac"] {
        no(name);
        yes(&format!("{name} = 0"));
        yes(&format!("Not {name}"));
        yes(&format!("{name} = False"));
        yes(&format!("{name} & \"a\" = \"0a\""));
        yes(&format!("{name} + 1 = 1"));
    }
}

#[test]
fn the_platform_names_ignore_case() {
    yes("WIN32 = 1");
    yes("win32 = 1");
    no("WIN16");
    no("mac");
}

#[test]
fn vba6_is_not_defined_for_a_project_module() {
    no("VBA6");
    yes("VBA6 = 0");
}

#[test]
fn a_name_with_no_definition_is_the_integer_zero() {
    no("Zork = \"\"");
    yes("Zork = 0");
    yes("(Zork & \"a\") = \"0a\"");
    yes("Zork + 1 = 1");
    yes("Not Zork");
    no("Zork");
    yes("Zork <> \"\"");
}

const TYPE_MISMATCH: u32 = 13;

fn yes_with(preamble: &str, cond: &str) {
    assert_eq!(run(preamble, cond), ("Yes".to_string(), vec![]), "{preamble}{cond} should be true");
}
fn no_with(preamble: &str, cond: &str) {
    assert_eq!(run(preamble, cond), ("No".to_string(), vec![]), "{preamble}{cond} should be false");
}

#[test]
fn a_number_reached_by_name_is_compared_with_text_as_text() {
    let five = "#Const X = 5\n";
    no_with(five, "X < \"10\"");
    yes_with(five, "X > \"10\"");
    no_with(five, "\"10\" > X");
    no_with(five, "X = \"5.0\"");
    yes_with(five, "X = \"5\"");
    no_with(five, "X = \"a\"");
    yes_with(five, "X < \"a\"");
    no_with(five, "X + 1 = \"6.0\"");
    yes_with(five, "X + 1 = \"6\"");
    yes_with("#Const X = -5\n", "X = \"-5\"");
    // Text is compared without regard to case.
    yes_with("#Const S = \"ABC\"\n#Const X = 1\n", "S = \"abc\" And X = \"1\"");
    // A name anywhere in an operand is enough.
    no_with(five, "(X * 1) = \"5.0\"");
    no_with(five, "CInt(X) = \"5.0\"");
}

#[test]
fn a_number_written_out_is_compared_with_text_as_a_number() {
    yes("5 = \"5.0\"");
    yes("5 < \"10\"");
    yes("1 + 1 = \"2.0\"");
    yes("CInt(5) = \"5.0\"");
    for cond in ["0 = \"\"", "0 = \"a\"", "(0) = \"\"", "1 + 1 = \"\"", "CInt(0) = \"\"", "1.5 < \"a\""] {
        assert_eq!(run("", cond), ("No".to_string(), vec![(TYPE_MISMATCH, 0, 3)]), "{cond}");
    }
}

#[test]
fn other_values_reached_by_name_are_compared_with_text_by_their_own_text() {
    yes_with("#Const X = True\n", "X = \"True\"");
    no_with("#Const X = True\n", "X = \"-1\"");
    yes_with("#Const X = Empty\n", "X = \"\"");
    no_with("#Const X = Empty\n", "X = \"0\"");
    yes_with("#Const X = 1.5@\n", "X = \"1.5\"");
    no_with("#Const X = 1.5\n", "X = \"1,5\"");
    yes_with("#Const X = 1E20\n", "X = \"1E+20\"");
    // Written out, the text is read as a boolean or a number.
    yes("True = \"True\"");
    yes("True = \"-1\"");
}

#[test]
fn a_boolean_written_out_is_compared_with_text_read_as_a_boolean() {
    for cond in [
        "True = \"1\"", "True = \"2\"", "False = \"0\"", "False = \"False\"", "True = \"true\"", "\"True\" = True",
        "\"-1\" = True", "True > \"False\" = False", "False < \"True\" = False",
    ] {
        yes(cond);
    }
    no("True = \"0\"");
    no("False = \"1\"");
    assert_eq!(run("", "True = \"yes\""), ("No".to_string(), vec![(TYPE_MISMATCH, 0, 3)]));
}

#[test]
fn a_name_in_an_arithmetic_operand_still_needs_a_number() {
    for cond in ["X + S = 1", "S + 1 = 1", "-S = 1", "CInt(S) = 1", "X * S = 1", "X Mod S = 1", "E + 1 = 1"] {
        let pre = "#Const S = \"a\"\n#Const X = 5\n#Const E = \"\"\n";
        let src = format!("{pre}#If {cond} Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n");
        let p = parse(&src);
        assert_eq!(p.found.first().map(|f| f.0), Some(TYPE_MISMATCH), "{cond}");
    }
}

#[test]
fn a_defined_constant_replaces_the_predefined_value() {
    assert_eq!(run("#Const Win32 = 5\n", "Win32 = 5"), ("Yes".to_string(), vec![]));
    assert_eq!(run("#Const Mac = 1\n", "Mac"), ("Yes".to_string(), vec![]));
    assert_eq!(run("#Const Win16 = \"x\"\n", "Win16 = \"x\""), ("Yes".to_string(), vec![]));
}

#[test]
fn a_project_argument_replaces_the_predefined_value() {
    let src = b"#If Win32 = 0 And Mac = 7 Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n";
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src);
    parser.set_conditional_constants(vec![
        ("win32".to_string(), CcValue::Integer(0)),
        ("mac".to_string(), CcValue::Integer(7)),
    ]);
    let top = parser.parse_module(&mut arena);
    assert!(parser.diagnostics.items().is_empty());
    drop(parser);
    let ExprNode::ProcDecl { name, .. } = arena.get(top[0]) else { panic!("no procedure") };
    assert_eq!(ctx.sym_name(*name as usize), "Yes");
}
