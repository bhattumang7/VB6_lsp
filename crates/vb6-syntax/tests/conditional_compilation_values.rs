//! Values and operators of conditional-compilation expressions, directive
//! spellings and line ends, and the project's conditional arguments.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprArena, ExprNode};
use vb6_syntax::frontend::cond_compile::{parse_cond_comp_args, CcValue};
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

const EXPECTED: u32 = 0x9c6f;

/// The name of the procedure that survives `#If <cond> Then Sub Yes #Else Sub No`.
fn taken(preamble: &str, cond: &str) -> String {
    let src = format!("{preamble}#If {cond} Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n");
    let p = parse(&src);
    assert_eq!(p.found, vec![], "{cond}");
    assert_eq!(p.top.len(), 1, "{cond}");
    let ExprNode::ProcDecl { name, .. } = p.node(p.top[0]) else { panic!("not a procedure") };
    p.name(*name).to_string()
}

fn yes(cond: &str) {
    assert_eq!(taken("", cond), "Yes", "{cond} should be true");
}
fn no(cond: &str) {
    assert_eq!(taken("", cond), "No", "{cond} should be false");
}

#[test]
fn floating_point_literals_and_arithmetic() {
    yes("1.5");
    yes("0.4");
    yes("2.5 > 2");
    yes("1 / 2 = 0.5");
    yes("7 / 2 = 3.5");
    yes("3 / 2 = 1.5");
    no("7 / 2 = 3");
    yes("1E2 = 100");
    yes("2147483648 > 2147483647");
    yes("10 \\ 3 = 3");
    yes("7.5 Mod 2 = 0");
    yes("5 Mod 3 = 2");
}

#[test]
fn power_and_concatenation() {
    yes("2 ^ 3 = 8");
    no("2 ^ 3 = 9");
    yes("2 * 3 ^ 2 = 18");
    yes("-2 ^ 2 = -4");
    yes("2 ^ 3 ^ 2 = 64");
    yes("\"a\" & \"b\" = \"ab\"");
    yes("1 & 2 = \"12\"");
    yes("1 + 2 & 3 = \"33\"");
}

#[test]
fn string_comparisons() {
    yes("\"a\" < \"b\"");
    no("\"b\" < \"a\"");
    yes("\"b\" > \"a\"");
    yes("\"abc\" = \"abc\"");
    no("\"A\" = \"a\"");
    yes("\"1\" = 1");
}

#[test]
fn not_is_bitwise() {
    yes("Not 5 = -6");
    yes("Not 0");
    yes("Not False");
    no("Not True");
    yes("Not (1 = 2)");
}

#[test]
fn bitwise_operators_act_on_whole_numbers() {
    yes("(6 And 3) = 2");
    yes("(6 Or 3) = 7");
    yes("(6 Xor 3) = 5");
    yes("(0 Eqv 0) = -1");
    yes("(0 Imp 0) = -1");
}

#[test]
fn constants_keep_their_types() {
    assert_eq!(taken("#Const D = 2.5\n", "D > 2"), "Yes");
    assert_eq!(taken("#Const S = \"x\"\n", "S = \"x\""), "Yes");
    assert_eq!(taken("#Const A = 3 / 2\n", "A = 1.5"), "Yes");
    assert_eq!(taken("#Const A = 1\n#Const B = A + 1\n", "B = 2"), "Yes");
    assert_eq!(taken("", "Undefined = 0"), "Yes");
    assert_eq!(taken("#Const verbose = 1\n", "VERBOSE"), "Yes");
}

#[test]
fn one_word_end_if_and_blank_after_the_hash() {
    for src in [
        "#If True Then\nSub A()\nEnd Sub\n#EndIf\n",
        "# If True Then\nSub A()\nEnd Sub\n# End If\n",
        "#  If True Then\nSub A()\nEnd Sub\n#   EndIf\n",
        "# Const K = 1\n#If K Then\nSub A()\nEnd Sub\n#End If\n",
    ] {
        let p = parse(src);
        assert_eq!(p.found, vec![], "{src:?}");
        assert_eq!(p.top.len(), 1, "{src:?}");
    }
}

#[test]
fn a_colon_after_a_directive_is_a_syntax_error() {
    for src in [
        "#Const A = 1 : #If A Then
Sub Yes()
End Sub
#End If
",
        "#Const A = 1:
",
        "#If 1 Then:
Sub A()
End Sub
#End If
",
        "#If 1 Then
Sub A()
End Sub
#Else:
#End If
",
    ] {
        let found = parse(src).found;
        assert_eq!(found[0].0, EXPECTED, "{src:?}");
        assert_eq!(found[0].1, offset_of(src, ":", 0), "{src:?}");
    }
}

#[test]
fn rem_after_a_directive_is_a_syntax_error_but_an_apostrophe_is_not() {
    let src = "#Const A = 1 Rem c
";
    // The comment word takes the rest of the line with it.
    assert_eq!(parse(src).found[0], (EXPECTED, offset_of(src, "Rem", 0), 5));
    assert_eq!(parse("#Const A = 1 'c
").found, vec![]);
}

#[test]
fn a_constant_name_is_not_a_reserved_word() {
    for word in ["Debug", "Date", "Me", "String", "Long", "If", "Rem", "Attribute", "VB_Name"] {
        let src = format!("#Const {word} = 1
");
        let found = parse(&src).found;
        let len = if word == "Rem" { "Rem = 1".len() } else { word.len() };
        assert_eq!(found[0], (EXPECTED, offset_of(&src, word, 0), len as u32), "{word}");
    }
    for word in ["Text", "Name", "Error", "Line", "Mid"] {
        assert_eq!(parse(&format!("#Const {word} = 1
")).found, vec![], "{word}");
    }
}

#[test]
fn an_elseif_after_a_taken_branch_is_still_read() {
    let src = "#If 1 Then
#ElseIf
#End If
";
    assert_eq!(parse(src).found[0].1, offset_of(src, "
#End", 0));
    let src = "#If 1 Then
#ElseIf 1
#End If
";
    assert_eq!(parse(src).found[0].0, 0x9c70_u32.min(parse(src).found[0].0));
}

#[test]
fn if_needs_then_and_nothing_follows_a_directive() {
    let src = "#If 1\nSub A()\nEnd Sub\n#End If\n";
    let found = parse(src).found;
    assert_eq!(found[0].0, EXPECTED);
    assert_eq!(found[0].1, offset_of(src, "\nSub", 0));

    let src = "#If 1 Then foo\nSub A()\nEnd Sub\n#End If\n";
    assert_eq!(parse(src).found, vec![(EXPECTED, offset_of(src, "foo", 0), 3)]);

    let src = "#If 1 Then\nSub A()\nEnd Sub\n#End If foo\n";
    assert_eq!(parse(src).found, vec![(EXPECTED, offset_of(src, "foo", 0), 3)]);

    let src = "#If 1 Then\nSub A()\nEnd Sub\n#Else foo\n#End If\n";
    assert_eq!(parse(src).found, vec![(EXPECTED, offset_of(src, "foo", 0), 3)]);
}

#[test]
fn a_missing_expression_does_not_swallow_the_next_line() {
    let src = "#Const A =\nSub F()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(EXPECTED, offset_of(src, "\nSub", 0), 0)]);
    assert_eq!(p.top.len(), 1);
}

#[test]
fn const_name_must_be_a_name() {
    let src = "#Const 5 = 1\n";
    let found = parse(src).found;
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].1, offset_of(src, "5", 0));
}

#[test]
fn project_arguments_are_visible_to_the_module() {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let src = b"#If MyFlag And Name = \"x\" Then\nSub Yes()\nEnd Sub\n#Else\nSub No()\nEnd Sub\n#End If\n";
    let mut parser = Parser::new(&mut ctx, src);
    parser.set_conditional_constants(parse_cond_comp_args("MyFlag = 1 : Name = \"x\""));
    let top = parser.parse_module(&mut arena);
    assert!(parser.diagnostics.items().is_empty());
    drop(parser);
    let ExprNode::ProcDecl { name, .. } = arena.get(top[0]) else { panic!("no procedure") };
    assert_eq!(ctx.sym_name(*name as usize), "Yes");
}

#[test]
fn argument_text_is_split_and_typed() {
    assert_eq!(
        parse_cond_comp_args("A = 1 : B = \"x:y\" : C=2.5:D : E = \"say \"\"hi\"\"\" : F = -3"),
        vec![
            ("a".to_string(), CcValue::Int(1)),
            ("b".to_string(), CcValue::Str("x:y".to_string())),
            ("c".to_string(), CcValue::Double(2.5)),
            ("d".to_string(), CcValue::Empty),
            ("e".to_string(), CcValue::Str("say \"hi\"".to_string())),
            ("f".to_string(), CcValue::Int(-3)),
        ]
    );
    assert_eq!(parse_cond_comp_args(""), vec![]);
}
