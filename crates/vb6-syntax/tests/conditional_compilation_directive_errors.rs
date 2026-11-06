//! What a conditional-compilation line reports, and where: lines that do not read as a
//! directive are ordinary lines, the errors of the directive lines come before every other
//! error of the module, and a directive word inside a continued line is no directive.

use vb6_syntax::frontend::ast::ExprArena;
use vb6_syntax::frontend::parser::Parser;
use vb6_syntax::frontend::scanner::ScannerContext;

const ORPHAN: &str = "An # ElseIf, # Else, or # EndIf must be preceded by an # If clause";
const UNTERMINATED: &str = "You must terminate the # If block with an # EndIf";
const ELSE_AFTER_ELSE: &str = "An # Else clause must be preceded by an # If";
const SYNTAX: &str = "Syntax error";

/// Every diagnostic as `(message, 1-based line)`, in report order.
fn report(src: &str) -> Vec<(String, usize)> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    parser.parse_module(&mut arena);
    parser
        .diagnostics
        .items()
        .iter()
        .map(|d| {
            let line = 1 + src.as_bytes()[..d.span.start as usize].iter().filter(|b| **b == b'\n').count();
            (d.message().expect("a message"), line)
        })
        .collect()
}

fn first(src: &str) -> (String, usize) {
    report(src).into_iter().next().unwrap_or_else(|| panic!("no diagnostic for {src:?}"))
}

fn expect(src: &str, message: &str, line: usize) {
    assert_eq!(first(src), (message.to_string(), line), "{src:?}");
}

// ── directive words inside a continued line ──────────────────────────────────

#[test]
fn a_directive_word_after_a_continuation_is_no_directive() {
    // The `#If` is part of the statement; the `#End If` that follows it begins a line and is
    // a directive with nothing to close.
    expect("Sub A()\nDim x As Long\nx = 1 + _\n#If True Then\n2\n#End If\nEnd Sub\n", ORPHAN, 6);
    expect("Sub A()\nDim x As Long\nx = 1 + _\n#If False Then\n2\n#End If\n3\nEnd Sub\n", ORPHAN, 6);
}

#[test]
fn an_end_if_inside_a_continued_line_does_not_close_the_block() {
    expect("Sub A()\nDim x As Long\n#If True Then\nx = 1 + _\n#End If\n2\nEnd Sub\n", UNTERMINATED, 7);
}

#[test]
fn a_syntax_error_in_a_continued_line_is_on_its_first_line() {
    expect("Sub A()\nDim x As Long\nx = 1 + _\n2 + _\n#If False Then\n3 + _\n#End If\n4\nEnd Sub\n", SYNTAX, 3);
    expect("Sub A()\nDim x As Long\nx = 1 + _\n#Const Z = 1\n2\nEnd Sub\n", SYNTAX, 3);
}

#[test]
fn a_directive_word_inside_a_continued_line_defines_nothing() {
    let src = "#Const Z = 1\nSub A()\nDim x As Long\nx = 1 + _\n#Const Z = 2\nEnd Sub\n";
    // No duplicate definition: the second `#Const` is no directive.
    let found = report(src);
    assert_eq!(found.first(), Some(&(SYNTAX.to_string(), 4)), "{found:?}");
    assert!(found.iter().all(|(m, _)| m != "Duplicate definition"), "{found:?}");
}

// ── lines that do not read as a directive ────────────────────────────────────

#[test]
fn an_if_line_that_does_not_read_through_is_no_directive() {
    // Its `#End If` has nothing to close, and is reported before the line itself.
    expect("#Const X = 1\n#If X\nConst Q = 1\n#End If\n", ORPHAN, 4);
    expect("#If Then\n#End If\n", ORPHAN, 2);
    expect("#Const X = 1\n#If X Then Dim\nConst Q = 1\n#End If\n", ORPHAN, 4);
    expect("#If (1 Then\n#End If\n", ORPHAN, 2);
    expect("#If 1 + Then\n#End If\n", ORPHAN, 2);
    expect("#If Foo( Then\n#End If\n", ORPHAN, 2);
    expect("Sub M()\nDim x As Long\n#If 1\nx = 1\n#End If\nEnd Sub\n", ORPHAN, 5);
    expect("Sub M()\nDim x As Long\n#If 1 Then x = 2\nx = 1\n#End If\nEnd Sub\n", ORPHAN, 5);
    expect("Sub M()\nDim x As Long\nx = 1: #If 1 Then\n#End If\nEnd Sub\n", ORPHAN, 4);
}

#[test]
fn an_end_if_line_that_does_not_read_through_leaves_the_block_open() {
    // The block is reported once, on the last line of the module.
    expect("#If 1 Then\n#End If x\nSub P()\nEnd Sub\n", UNTERMINATED, 4);
    expect("#If 1 Then\n#EndIf x\nSub P()\nEnd Sub\n", UNTERMINATED, 4);
    expect("#If 1 Then\n#End If :\nSub P()\nEnd Sub\n", UNTERMINATED, 4);
    expect("#If 1 Then\n#End\nSub P()\nEnd Sub\n", UNTERMINATED, 4);
    expect("#If 1 Then\n#End Sub\nSub P()\nEnd Sub\n", UNTERMINATED, 4);
    expect("Sub M()\nDim x As Long\n#If 1 Then\nx = 1\n#End If x\nEnd Sub\n", UNTERMINATED, 6);
    expect("Sub M()\nDim x As Long\n#If 1 Then\nx = 1\n#End If :\nEnd Sub\n", UNTERMINATED, 6);
    expect("Sub M()\n#If 1 Then\n#End Sub\nEnd Sub\n", UNTERMINATED, 4);
}

#[test]
fn a_bare_end_in_a_procedure_is_a_statement_that_fails() {
    expect("Sub M()\nDim x As Long\n#If 1 Then\nx = 1\n#End\n#End If\nEnd Sub\n", SYNTAX, 5);
}

#[test]
fn an_else_line_with_trailing_text_names_what_it_expected() {
    expect("#If 1 Then\n#Else x\n#End If\n", "Expected: end of statement", 2);
}

// ── order of the reports ─────────────────────────────────────────────────────

#[test]
fn the_errors_of_directive_lines_come_before_any_other_error() {
    expect("Sub S()\nx = = 1\nEnd Sub\n#End If\n", ORPHAN, 4);
    expect("Sub S()\nx = = 1\nEnd Sub\n#If True Then\n", UNTERMINATED, 4);
    expect("Sub S()\nx = = 1\nEnd Sub\n#If 1 / 0 Then\n#End If\n", "Division by zero", 4);
    expect("Sub S()\nx = = 1\nEnd Sub\n#If True Then\n#Else\n#Else\n#End If\n", ELSE_AFTER_ELSE, 6);
    expect("#Const A = 1\nSub S()\nx = = 1\nEnd Sub\n#Const A = 2\n", "Duplicate definition", 5);
}

#[test]
fn the_other_errors_follow_in_source_order() {
    let found = report("Sub S()\nx = = 1\nEnd Sub\n#End If\nSub T()\ny = = 2\nEnd Sub\n");
    assert_eq!(found[0], (ORPHAN.to_string(), 4));
    let lines: Vec<usize> = found[1..].iter().map(|f| f.1).collect();
    assert!(lines.windows(2).all(|w| w[0] <= w[1]), "{found:?}");
    assert_eq!(lines.first(), Some(&2), "{found:?}");
    assert_eq!(lines.last(), Some(&6), "{found:?}");
    assert!(lines.iter().all(|l| *l == 2 || *l == 6), "{found:?}");
}

#[test]
fn directive_lines_in_a_block_that_is_not_compiled_are_still_read() {
    // A line that is no directive inside code that is not compiled is ignored.
    assert_eq!(report("#If False Then\n#Foo\n#End If\n"), vec![]);
    expect("#If False Then\n#If 1 2 Then\n#End If\n#End If\n", ORPHAN, 4);
}

// ── what a failing directive expected ────────────────────────────────────────

#[test]
fn a_failing_condition_names_what_it_expected() {
    expect("#If True Then\n#ElseIf 1 2 Then\n#End If\n", "Expected: Then", 2);
    expect("#If True Then\n#ElseIf Foo Bar Then\n#End If\n", "Expected: Then", 2);
    expect("#If False Then\n#ElseIf True Then\n#ElseIf 1 2 Then\n#End If\n", "Expected: Then", 3);
    expect("#If 1 Then\n#ElseIf 1\n#End If\n", "Expected: Then", 2);
    expect("#If True Then\n#ElseIf Foo = Then\n#End If\n", "Expected: expression", 2);
}

#[test]
fn a_number_sign_that_names_no_directive() {
    let list = "Expected: If or Else or ElseIf or End or EndIf or Const";
    expect("#Foo\nSub P()\nEnd Sub\n", list, 1);
    expect("#\nSub P()\nEnd Sub\n", list, 1);
    expect("#1\nSub P()\nEnd Sub\n", list, 1);
    // Inside a procedure only the general error is given.
    expect("Sub M()\n#Foo\nEnd Sub\n", SYNTAX, 2);
}

#[test]
fn a_const_directive_names_what_it_expected() {
    expect("#Const X 1\nSub P()\nEnd Sub\n", "Expected: =", 1);
    expect("#Const X =\nSub P()\nEnd Sub\n", "Expected: expression", 1);
    expect("#Const 1 = 1\nSub P()\nEnd Sub\n", "Expected: identifier", 1);
    expect("Sub M()\n#Const X 1\nEnd Sub\n", SYNTAX, 2);
    expect("Sub M()\n#Const 1 = 1\nEnd Sub\n", SYNTAX, 2);
}

#[test]
fn the_block_pass_reports_one_open_block_once() {
    // Two blocks are open; the innermost is reported, once.
    let found = report("#If True Then\n#If True Then\nSub P()\nEnd Sub\n");
    assert_eq!(found, vec![(UNTERMINATED.to_string(), 4)]);
}
