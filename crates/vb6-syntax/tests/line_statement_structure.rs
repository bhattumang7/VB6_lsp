//! How a logical line divides into statements: where a line number may stand, what may
//! follow a jump written after `Then` or `Else`, the colon-led form of a single-line `If`,
//! the statements on an `ElseIf` line, `Go To`/`Go Sub`, and where an `Else` may end a statement.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, FileIoKind, LabelRef, NodeId};

const SYNTAX_ERROR: u32 = 0x9c70;
const EXPECTED: u32 = 0x9c6f;
const MUST_BE_FIRST_ON_LINE: u32 = 0x9c71;

fn label_text(p: &Parsed, target: &LabelRef) -> String {
    match target {
        LabelRef::Line(n) => n.to_string(),
        LabelRef::Name(sym) => p.name(*sym).to_string(),
    }
}

/// A one-line rendering of a node, exact for every node kind these tests produce.
fn show(p: &Parsed, id: NodeId) -> String {
    let list = |ids: &[NodeId]| ids.iter().map(|n| show(p, *n)).collect::<Vec<_>>().join("; ");
    match p.node(id) {
        ExprNode::NameRef { sym, .. } => p.name(*sym).to_string(),
        ExprNode::Literal { lit: AstLit::Int(n) } => n.to_string(),
        ExprNode::Assign { target, value } => format!("{} = {}", show(p, *target), show(p, *value)),
        ExprNode::GoTo { target } => format!("GoTo {}", label_text(p, target)),
        ExprNode::GoSub { target } => format!("GoSub {}", label_text(p, target)),
        ExprNode::Label { target } => format!("Label {}", label_text(p, target)),
        ExprNode::Block { stmts } => format!("[{}]", list(stmts)),
        ExprNode::If { cond, then_body, else_body } => format!(
            "If {} Then {} Else {}",
            show(p, *cond),
            show(p, *then_body),
            else_body.map_or("-".to_string(), |e| show(p, e))
        ),
        ExprNode::CallStmt { callee, args } => format!("Call {}({})", show(p, *callee), show(p, *args)),
        ExprNode::ArgList { args } => list(args).replace("; ", ", "),
        ExprNode::Paren { inner } => format!("({})", show(p, *inner)),
        ExprNode::Stop => "Stop".to_string(),
        ExprNode::FileIoStmt { kind: FileIoKind::Close, channel: None, args, channel_marks: _, open_mode: 0, open_flags: 0 } => {
            format!("Close {}", list(args).replace("; ", ", "))
        }
        other => panic!("node not covered by the test renderer: {}", std::mem::discriminant(other) == std::mem::discriminant(&ExprNode::Stop)),
    }
}

/// The statements of the only procedure of `Sub S() ... End Sub` around `body`, rendered.
fn statements(body: &str) -> Vec<String> {
    let p = parse(&format!("Sub S()\n{body}\nEnd Sub\n"));
    assert_eq!(p.found, vec![], "for {body:?}");
    p.proc_body(0).iter().map(|s| show(&p, *s)).collect()
}

/// The diagnostics of `Sub S() ... End Sub` around `body`, as (code, start, length) in `src`.
fn diagnostics(body: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let found = parse(&src).found;
    (src, found)
}

// ── Line numbers ────────────────────────────────────────────────────────────────────────────

#[test]
fn a_line_number_starts_a_line_and_may_be_followed_by_a_label_or_a_statement() {
    assert_eq!(statements("10 Lbl: i = 1"), ["Label 10", "Label Lbl", "i = 1"]);
    assert_eq!(statements("10: i = 1"), ["Label 10", "i = 1"]);
    assert_eq!(statements("10 i = 1"), ["Label 10", "i = 1"]);
    assert_eq!(statements("10"), ["Label 10"]);
}

#[test]
fn a_second_line_number_after_a_numbered_label_is_a_syntax_error() {
    for body in ["10: 20 i = 1", "10 : 20 i = 1", "10:20 i = 1", "10: 20: i = 1", "10: 20"] {
        let (src, found) = diagnostics(body);
        assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "20", 0), 2)], "for {body:?}");
    }
}

#[test]
fn a_line_number_after_a_name_label_is_a_syntax_error() {
    let (src, found) = diagnostics("Lbl: 20 i = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "20", 0), 2)]);
    let (src, found) = diagnostics("10 Lbl: 20 i = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "20", 0), 2)]);
    // A hexadecimal number is a number too.
    let (src, found) = diagnostics("L1: &H10");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "&H10", 0), 4)]);
    let (src, found) = diagnostics("10 L1: &H10");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "&H10", 0), 4)]);
}

#[test]
fn a_line_number_after_a_statement_is_a_syntax_error() {
    let (src, found) = diagnostics("i = 1: 10 j = 2");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "10", 0), 2)]);
}

#[test]
fn a_line_number_straight_after_a_line_number_is_a_syntax_error() {
    let (src, found) = diagnostics("10 20 i = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "20", 0), 2)]);
}

// ── A jump written after Then or Else ───────────────────────────────────────────────────────

#[test]
fn a_line_number_after_then_or_else_is_a_jump() {
    assert_eq!(
        statements("If c Then 10 Else 20\n10: i = 1\n20: i = 2"),
        ["If c Then [GoTo 10] Else [GoTo 20]", "Label 10", "i = 1", "Label 20", "i = 2"]
    );
    assert_eq!(
        statements("If c Then 10: j = 1\n10: j = 2"),
        ["If c Then [GoTo 10; j = 1] Else -", "Label 10", "j = 2"]
    );
    assert_eq!(
        statements("If c Then i = 1 Else 20: j = 2\n20: j = 3"),
        ["If c Then [i = 1] Else [GoTo 20; j = 2]", "Label 20", "j = 3"]
    );
    assert_eq!(
        statements("If c Then i = 1: Else 20\n20: j = 3"),
        ["If c Then [i = 1] Else [GoTo 20]", "Label 20", "j = 3"]
    );
}

#[test]
fn a_jump_after_then_or_else_is_followed_by_a_separator() {
    let (src, found) = diagnostics("If c Then 10 i = 2\n10: i = 1");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "i = 2", 0), 1)]);
    let (src, found) = diagnostics("If c Then i = 1 Else 20 i = 3\n20: i = 1");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "i = 3", 0), 1)]);
    let (src, found) = diagnostics("If c Then 10 20\n10: i = 1\n20: i = 1");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "20", 0), 2)]);
    let (src, found) = diagnostics("If c Then 10 Else 20 Else 30\n10: i = 1\n20: i = 1");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Else 30", 0), 4)]);
}

// ── The colon-led single-line If ────────────────────────────────────────────────────────────

#[test]
fn then_followed_by_a_colon_reads_the_line_as_statements() {
    assert_eq!(statements("If c Then: i = 2"), ["If c Then [i = 2] Else -"]);
    assert_eq!(statements("If c Then: i = 2: j = 3"), ["If c Then [i = 2; j = 3] Else -"]);
    assert_eq!(statements("If c Then:: i = 2"), ["If c Then [i = 2] Else -"]);
    assert_eq!(statements("If c Then:"), ["If c Then [] Else -"]);
}

#[test]
fn an_else_follows_a_colon_led_statement_only_after_a_colon() {
    assert_eq!(statements("If c Then: i = 2: Else j = 3"), ["If c Then [i = 2] Else [j = 3]"]);
    assert_eq!(statements("If c Then: Else j = 3"), ["If c Then [] Else [j = 3]"]);
    assert_eq!(statements("If c Then i = 2: Else j = 3"), ["If c Then [i = 2] Else [j = 3]"]);
    let (src, found) = diagnostics("If c Then: i = 2 Else j = 3");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Else", 0), 4)]);
    let (src, found) = diagnostics("If c Then: GoTo L Else j = 3\nL:");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Else", 0), 4)]);
    let (src, found) = diagnostics("If c Then: Exit Sub Else j = 3");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Else", 0), 4)]);
}

#[test]
fn a_line_number_in_a_colon_led_if_is_no_jump() {
    let (src, found) = diagnostics("If c Then: 10 i = 2\n10: i = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "10 i", 0), 2)]);
    let (src, found) = diagnostics("If c Then: Else 20\n20: i = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "20", 0), 2)]);
    let (src, found) = diagnostics("If c Then: i = 1: Else 20\n20: i = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "20", 0), 2)]);
    let (src, found) = diagnostics("If c Then i = 1 Else: 20\n20: i = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "20", 0), 2)]);
}

#[test]
fn a_colon_led_if_must_be_first_on_its_line() {
    let (src, found) = diagnostics("i = 1: If c Then: i = 2");
    assert_eq!(found, vec![(MUST_BE_FIRST_ON_LINE, offset_of(&src, "If", 0), 2)]);
    let (src, found) = diagnostics("i = 1: If c Then: Else i = 3");
    assert_eq!(found, vec![(MUST_BE_FIRST_ON_LINE, offset_of(&src, "If", 0), 2)]);
    // The plain single-line form may follow other statements.
    assert_eq!(
        statements("i = 1: If c Then i = 2 Else i = 3"),
        ["i = 1", "If c Then [i = 2] Else [i = 3]"]
    );
}

#[test]
fn nested_single_line_ifs_share_the_else() {
    assert_eq!(
        statements("If p Then If q Then x = 1 Else x = 2 Else x = 3"),
        ["If p Then [If q Then [x = 1] Else [x = 2]] Else [x = 3]"]
    );
    assert_eq!(
        statements("If p Then If q Then x = 1: Else x = 2"),
        ["If p Then [If q Then [x = 1] Else [x = 2]] Else -"]
    );
    assert_eq!(
        statements("If p Then: If q Then x = 1 Else x = 2"),
        ["If p Then [If q Then [x = 1] Else [x = 2]] Else -"]
    );
    let (src, found) = diagnostics("If p Then If q Then: x = 1 Else x = 2");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Else", 0), 4)]);
    let (src, found) = diagnostics("If p Then If q Then x = 1 Else x = 2 Else x = 3 Else x = 4");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Else x = 4", 0), 4)]);
    let (src, found) = diagnostics("If p Then x = 1 Else x = 2 Else x = 3");
    assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Else x = 3", 0), 4)]);
}

#[test]
fn a_second_else_directly_after_else_is_a_syntax_error() {
    let (src, found) = diagnostics("If c Then: Else Else i = 3");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "Else Else", 0) + 5, 4)]);
}

// ── Go To / Go Sub / GoTo in an If ──────────────────────────────────────────────────────────

#[test]
fn go_is_a_jump_only_with_to_or_sub() {
    assert_eq!(statements("Go To L\nL:"), ["GoTo L", "Label L"]);
    assert_eq!(statements("Go Sub L\nL:"), ["GoSub L", "Label L"]);
    assert_eq!(statements("Go To 10\n10: i = 1"), ["GoTo 10", "Label 10", "i = 1"]);
    for body in ["Go L\nL:", "Go 1\n1: i = 1", "Go P\nP:"] {
        let (src, found) = diagnostics(body);
        let at = offset_of(&src, "Go ", 0) + 3;
        assert_eq!(found, vec![(SYNTAX_ERROR, at, 1)], "for {body:?}");
    }
    let (src, found) = diagnostics("Go");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "Go", 0) + 2, 0)]);
    // `Go` followed by an operator is a name.
    assert_eq!(statements("Go = 1"), ["Go = 1"]);
}

#[test]
fn an_if_without_then_takes_goto_but_not_go_to() {
    assert_eq!(statements("If c GoTo L\nL:"), ["If c Then [GoTo L] Else -", "Label L"]);
    assert_eq!(
        statements("If c GoTo L Else GoTo M\nL:\nM:"),
        ["If c Then [GoTo L] Else [GoTo M]", "Label L", "Label M"]
    );
    assert_eq!(statements("If c Then Go To L\nL:"), ["If c Then [GoTo L] Else -", "Label L"]);
    assert_eq!(statements("If c GoTo L: i = 1\nL:"), ["If c Then [GoTo L; i = 1] Else -", "Label L"]);
    for body in ["If c Go To L\nL:", "If c  Go To L\nL:", "If c Go To 10\n10: i = 1"] {
        let (src, found) = diagnostics(body);
        assert_eq!(found, vec![(EXPECTED, offset_of(&src, "Go", 0), 2)], "for {body:?}");
    }
}

// ── ElseIf lines ────────────────────────────────────────────────────────────────────────────

#[test]
fn statements_on_an_elseif_line_open_its_branch() {
    assert_eq!(
        statements("If p Then\nElseIf q Then x = 1\ny = 2\nEnd If"),
        ["If p Then [] Else If q Then [x = 1; y = 2] Else -"]
    );
    assert_eq!(
        statements("If p Then\nElseIf q Then 10\n10: x = 1\nEnd If"),
        ["If p Then [] Else If q Then [GoTo 10; Label 10; x = 1] Else -"]
    );
    assert_eq!(
        statements("If p Then\nElseIf q Then: x = 1\nElse\ny = 2\nEnd If"),
        ["If p Then [] Else If q Then [x = 1] Else [y = 2]"]
    );
    assert_eq!(
        statements("If p Then\nElseIf q Then x = 1: z = 3\nEnd If"),
        ["If p Then [] Else If q Then [x = 1; z = 3] Else -"]
    );
}

#[test]
fn an_else_on_an_elseif_line_is_an_error() {
    for (body, code) in [
        ("If p Then\nElseIf q Then x = 1 Else x = 2\nEnd If", EXPECTED),
        ("If p Then\nElseIf q Then x = 1: Else x = 2\nEnd If", SYNTAX_ERROR),
        ("If p Then\nElseIf q Then: x = 1: Else x = 2\nEnd If", SYNTAX_ERROR),
        ("If p Then\nElseIf q Then 10 Else x = 2\nEnd If\n10: x = 1", EXPECTED),
    ] {
        let (src, found) = diagnostics(body);
        // The first `Else` in the text is the start of `ElseIf`.
        assert_eq!(found, vec![(code, offset_of(&src, "Else", 1), 4)], "for {body:?}");
    }
}

#[test]
fn a_block_if_cannot_open_on_an_elseif_line() {
    let (src, found) = diagnostics("If p Then\nElseIf q Then If c Then\nEnd If\nEnd If");
    // The error stands at the line's end, behind the `Then` of the inner `If`.
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "Then\nEnd If\nEnd If", 0) + 4, 0)]);
}

// ── Close before Else, a trailing comma after a parenthesised first argument ──────────────────

#[test]
fn close_needs_a_channel_before_else() {
    assert_eq!(
        statements("If c Then Close #1 Else x = 1"),
        ["If c Then [Close 1] Else [x = 1]"]
    );
    assert_eq!(statements("If c Then Close : x = 1"), ["If c Then [Close ; x = 1] Else -"]);
    let (src, found) = diagnostics("If c Then Close Else x = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "Else", 0), 4)]);
    let (src, found) = diagnostics("Close Else");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "Else", 0), 4)]);
    // A comma needs a channel behind it, before a colon or an `Else` as before the line's end.
    let (src, found) = diagnostics("Close #1, : x = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, ":", 0), 1)]);
    let (src, found) = diagnostics("If c Then Close #1, Else x = 1");
    assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, "Else", 0), 4)]);
    let (src, found) = diagnostics("Close #1,");
    assert_eq!(found, vec![(SYNTAX_ERROR, src.len() as u32 - "\nEnd Sub\n".len() as u32, 0)]);
}

#[test]
fn a_call_with_a_parenthesised_first_argument_needs_an_argument_after_a_comma() {
    assert_eq!(statements("Foo (1), 2"), ["Call Foo((1), 2)"]);
    for (body, after, len) in [
        ("Foo (1),", "Foo (1),", 0),
        ("Foo (1), : x = 2", "Foo (1), ", 1),
        ("Foo (1), Else", "Foo (1), ", 4),
        ("If c Then Foo (1), Else x = 2", "If c Then Foo (1), ", 4),
    ] {
        let (src, found) = diagnostics(body);
        assert_eq!(found, vec![(SYNTAX_ERROR, offset_of(&src, after, 0) + after.len() as u32, len)], "for {body:?}");
    }
}

// ── Text after the closing line of a procedure ──────────────────────────────────────────────

#[test]
fn text_after_end_sub_on_its_line_is_a_syntax_error() {
    for (src, token, len) in [
        ("Sub S()\nEnd Sub x\n", "x", 1),
        ("Sub S()\nEnd Sub : Dim a\n", ":", 1),
        ("Sub S()\nEnd Sub :\n", ":", 1),
        ("Sub S()\nEnd Sub Rem done\n", "Rem", 8),
        ("Function F()\nEnd Function 1\n", "1", 1),
        ("Property Get P()\nEnd Property \"s\"\n", "\"s\"", 3),
    ] {
        let p = parse(src);
        assert_eq!(p.found, vec![(SYNTAX_ERROR, offset_of(src, token, 0), len)], "for {src:?}");
        assert_eq!(p.top.len(), 1, "for {src:?}");
    }
    for ok in ["Sub S()\nEnd Sub ' done\n", "Sub S()\nEnd Sub\n", "Sub S()\nEnd Sub"] {
        assert_eq!(parse(ok).found, vec![], "for {ok:?}");
    }
}

// ── Lines between the end of a procedure and the next one ───────────────────────────────────

const AFTER_END_OF_PROCEDURE: u32 = 0x9c91;

/// The diagnostics of `Sub S() / End Sub` followed by `line` on a line of its own.
fn after_end(line: &str) -> (String, Vec<(u32, u32, u32)>) {
    let src = format!("Sub S()\nEnd Sub\n{line}\n");
    let found = parse(&src).found;
    (src, found)
}

#[test]
fn a_line_that_reads_as_a_statement_is_text_after_the_end_of_the_procedure() {
    for (line, first) in [
        ("g = 1", "g"),
        ("g", "g"),
        ("10 g = 1", "10"),
        ("L1:", "L1"),
        ("L1: g = 1", "L1"),
        (":", ":"),
        (".x", "."),
        ("Else", "Else"),
        ("Else g = 1", "Else"),
        ("ElseIf x Then g = 1", "ElseIf"),
        ("End If", "End"),
        ("EndIf", "EndIf"),
        ("End With", "End"),
        ("End Select", "End"),
        ("End Sub", "End"),
        ("End Function", "End"),
        ("End Type", "End"),
        ("Next", "Next"),
        ("Next i, j", "Next"),
        ("Case 1", "Case"),
        ("Case Else", "Case"),
        ("Loop", "Loop"),
        ("Loop While x", "Loop"),
        ("Wend", "Wend"),
        ("Exit Sub", "Exit"),
        ("Exit Do", "Exit"),
        ("Print 1", "Print"),
        ("GoTo L1", "GoTo"),
        ("ReDim a(1)", "ReDim"),
        ("Static a", "Static"),
        ("Dim a", "Dim"),
        ("Const c = 1", "Const"),
        ("Call g", "Call"),
        ("x = 1: y = 2", "x"),
        ("1", "1"),
        ("Me.x", "Me"),
        ("Set a = b", "Set"),
        ("Resume Next", "Resume"),
        ("On Error GoTo 0", "On"),
        ("Debug.Print 1", "Debug"),
        ("With x", "With"),
        ("If x Then", "If"),
        ("For i = 1 To 2", "For"),
        ("Do", "Do"),
        ("Select Case x", "Select"),
        ("Option Base 1", "Option"),
        ("DefInt A-Z", "DefInt"),
        ("Declare Sub Q Lib \"k\" ()", "Declare"),
        ("Public a As Long", "Public"),
        ("Type T", "Type"),
        ("Enum E", "Enum"),
    ] {
        let (src, found) = after_end(line);
        let at = src.rfind(&format!("\n{line}\n")).unwrap() as u32 + 1;
        assert_eq!(found, vec![(AFTER_END_OF_PROCEDURE, at, first.len() as u32)], "for {line:?}");
    }
}

#[test]
fn a_line_that_does_not_read_as_a_statement_keeps_its_own_error() {
    for (line, code, at, len) in [
        ("Event E()", SYNTAX_ERROR, 0, 5),
        ("Public Event E()", SYNTAX_ERROR, 0, 6),
        ("Implements I", SYNTAX_ERROR, 0, 10),
        (",", SYNTAX_ERROR, 0, 1),
        ("&", SYNTAX_ERROR, 0, 1),
        ("#", SYNTAX_ERROR, 0, 1),
        ("; ", SYNTAX_ERROR, 0, 1),
        ("(g)", SYNTAX_ERROR, 0, 1),
        ("*", SYNTAX_ERROR, 0, 1),
        ("=", SYNTAX_ERROR, 0, 1),
        ("x = (1", EXPECTED, 6, 0),
        ("x = 1 +", EXPECTED, 7, 0),
        ("Do x", EXPECTED, 3, 1),
        ("Call", EXPECTED, 4, 0),
        ("Dim", EXPECTED, 3, 0),
        ("Const", EXPECTED, 5, 0),
        ("Option", EXPECTED, 6, 0),
        ("DefInt", EXPECTED, 6, 0),
        ("Type", EXPECTED, 4, 0),
        ("Enum", EXPECTED, 4, 0),
        ("Declare x", EXPECTED, 8, 1),
        ("End If x", EXPECTED, 7, 1),
        ("Wend x", EXPECTED, 5, 1),
        ("Loop Until", EXPECTED, 10, 0),
    ] {
        let (src, found) = after_end(line);
        let start = src.rfind(&format!("\n{line}\n")).unwrap() as u32 + 1;
        assert_eq!(found, vec![(code, start + at, len)], "for {line:?}");
    }
}

#[test]
fn every_statement_line_after_the_end_of_a_procedure_is_reported() {
    let src = "Sub S()\nEnd Sub\ng = 1\nh = 2\nSub T()\nEnd Sub\nk = 3\n";
    let found = parse(src).found;
    assert_eq!(
        found,
        vec![
            (AFTER_END_OF_PROCEDURE, offset_of(src, "g = 1", 0), 1),
            (AFTER_END_OF_PROCEDURE, offset_of(src, "h = 2", 0), 1),
            (AFTER_END_OF_PROCEDURE, offset_of(src, "k = 3", 0), 1),
        ]
    );
    // The same label on two such lines is no duplicate.
    let src = "Sub S()\nEnd Sub\nL1:\nL1:\n";
    assert_eq!(
        parse(src).found,
        vec![
            (AFTER_END_OF_PROCEDURE, offset_of(src, "L1", 0), 2),
            (AFTER_END_OF_PROCEDURE, offset_of(src, "L1", 1), 2),
        ]
    );
}

#[test]
fn comments_attributes_directives_and_procedures_may_follow_the_end_of_a_procedure() {
    for ok in [
        "' comment",
        "Rem comment",
        "Attribute VB_Description = \"d\"",
        "#Const k = 1",
        "#If True Then\n#End If",
        "Sub T()\nEnd Sub",
        "Private Function F()\nEnd Function",
        "Static Sub U()\nEnd Sub",
        "Property Get P()\nEnd Property",
    ] {
        let src = format!("Sub S()\nEnd Sub\n{ok}\n");
        assert_eq!(parse(&src).found, vec![], "for {ok:?}");
    }
}

#[test]
fn a_colon_only_line_before_the_first_procedure_is_blank() {
    assert_eq!(parse(":\nSub S()\nEnd Sub\n").found, vec![]);
}
