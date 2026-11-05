//! Conditional-compilation directives are resolved before the grammar: the
//! active branch parses in whatever construct surrounds the directives.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::{ExprNode, NodeId};

fn case_bodies(p: &common::Parsed, select: NodeId) -> Vec<usize> {
    let ExprNode::SelectCase { cases, .. } = p.node(select) else { panic!("SelectCase") };
    cases
        .iter()
        .map(|c| match p.node(*c) {
            ExprNode::CaseBlock { body, .. } | ExprNode::CaseElse { body } => p.block(*body).len(),
            _ => panic!("case"),
        })
        .collect()
}

const WRAPPED_CASE: &str = "Public Function Factor(ByVal s As String) As Integer\n    Select Case s\n        Case \"u\": Factor = -6\n        #If UNDEFINED = 0 Then\n        Case \"U\": Factor = -6\n        #End If\n        Case \"m\": Factor = -3\n    End Select\nEnd Function\n";

#[test]
fn case_clause_wrapped_in_true_directive() {
    let p = parse(WRAPPED_CASE);
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    assert_eq!(case_bodies(&p, body[0]), vec![1, 1, 1]);
}

#[test]
fn case_clause_wrapped_in_false_directive_is_dropped() {
    let p = parse(&WRAPPED_CASE.replace("UNDEFINED = 0", "UNDEFINED <> 0"));
    assert_eq!(p.diags, 0);
    assert_eq!(case_bodies(&p, p.proc_body(0)[0]), vec![1, 1]);
}

#[test]
fn directive_branches_choose_between_case_clauses() {
    let p = parse(
        "#Const MODE = 2\nSub A(x)\n Select Case x\n  #If MODE = 1 Then\n  Case 1: a = 1\n  #ElseIf MODE = 2 Then\n  Case 2: a = 2\n  Case 3: a = 3\n  #Else\n  Case Else: a = 4\n  #End If\n  Case 9: a = 9\n End Select\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    assert_eq!(case_bodies(&p, p.proc_body(0)[0]), vec![1, 1, 1]);
}

#[test]
fn nested_directives_and_inactive_garbage() {
    let p = parse(
        "Sub A()\n #If 0 Then\n  this is ( not valid\n  #If 1 Then\n  x = 1\n  #End If\n #Else\n  y = 1\n  #If 0 Then\n  z = 1\n  #Else\n  w = 1\n  #End If\n #End If\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    // Only `y = 1` and `w = 1` are active.
    let body = p.proc_body(0);
    assert_eq!(body.len(), 2);
    for n in &body {
        assert!(matches!(p.node(*n), ExprNode::Assign { .. }));
    }
}

#[test]
fn directives_may_split_an_if_else_block() {
    let p = parse(
        "Sub A()\n #If 1 Then\n If a Then\n #Else\n If b Then\n #End If\n  x = 1\n End If\nEnd Sub\n",
    );
    assert_eq!(p.diags, 0);
    let body = p.proc_body(0);
    assert_eq!(body.len(), 1);
    let ExprNode::If { then_body, else_body, .. } = p.node(body[0]) else { panic!() };
    assert_eq!(p.block(*then_body).len(), 1);
    assert!(else_body.is_none());
}

#[test]
fn module_level_directive_selects_procedures() {
    let p = parse("#If Win32 Then\nSub A()\nEnd Sub\n#Else\nSub B()\nEnd Sub\n#End If\nSub C()\nEnd Sub\n");
    assert_eq!(p.diags, 0);
    assert_eq!(p.top.len(), 2);
    let names: Vec<_> = p
        .top
        .iter()
        .map(|n| match p.node(*n) {
            ExprNode::ProcDecl { name, .. } => p.name(*name).to_string(),
            _ => panic!(),
        })
        .collect();
    assert_eq!(names, vec!["A", "C"]);
}

#[test]
fn const_inside_inactive_branch_is_still_defined() {
    let p = parse(
        "#If 0 Then\n#Const V = 2\n#End If\n#If V = 2 Then\nSub One()\nEnd Sub\n#Else\nSub Two()\nEnd Sub\n#End If\n",
    );
    assert_eq!(p.diags, 0);
    assert_eq!(p.top.len(), 1);
    let ExprNode::ProcDecl { name, .. } = p.node(p.top[0]) else { panic!() };
    assert_eq!(p.name(*name), "One");
}

#[test]
fn directive_with_trailing_comment_and_unterminated_block() {
    let src = "Sub A()\n #If 1 Then ' comment\n x = 1\n #End If ' done\n y = 2\nEnd Sub\n #If 0 Then\nSub Z()\n";
    let p = parse(src);
    // The unterminated `#If` is reported once, on the last line of the module.
    assert_eq!(p.found, vec![(0xdec0, offset_of(src, "Sub Z()", 0), 7)]);
    assert_eq!(p.proc_body(0).len(), 2);
}

#[test]
fn a_const_given_twice_is_a_duplicate_definition_at_the_second_directive() {
    for (src, second) in [
        ("#Const A = 1\n#Const a = 2\n", "#Const a"),
        ("#Const Win32 = 1\n#Const Win32 = 2\n", "#Const Win32 = 2"),
        ("#Const A = 1\n#If 0 Then\n#Const A = 2\n#End If\n", "#Const A = 2"),
    ] {
        let p = parse(src);
        assert_eq!(p.found, vec![(0xdec7, offset_of(src, second, 0), 6)], "{src:?}");
    }
    // A name given once, even one a project argument or predefined name already carries, is fine.
    assert_eq!(parse("#Const Win32 = 0\n").found, vec![]);
    assert_eq!(parse("#Const A = 1\n#Const B = 1\n").found, vec![]);
}

#[test]
fn a_malformed_directive_in_an_unused_branch_is_not_a_directive() {
    // The malformed inner `#If` opens nothing, so the second `#End If` is the orphan.
    let src = "#If False Then\n#If (1 Then\n#End If\n#End If\n";
    assert_eq!(parse(src).found, vec![(0xdebd, offset_of(src, "#End If", 1), 4)]);
    // An `#ElseIf` after a branch that was not taken is not read.
    assert_eq!(parse("#If False Then\n#ElseIf (( Then\n#End If\n").found, vec![]);
    // After a taken branch it is read, and its error counts.
    let src = "#If True Then\n#ElseIf (( Then\n#End If\n";
    let found = parse(src).found;
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].0, 0x9c6f);
}

#[test]
fn directives_in_unused_branches_are_still_evaluated() {
    let src = "#If False Then\n#If 1 / 0 Then\n#End If\n#End If\n";
    assert_eq!(parse(src).found, vec![(11, offset_of(src, "#If 1", 0), 3)]);
    let src = "#If False Then\n#Const A = 1 / 0\n#End If\n";
    assert_eq!(parse(src).found, vec![(11, offset_of(src, "#Const", 0), 6)]);
    let src = "#If True Then\n#Else\n#ElseIf 1 / 0 Then\n#End If\n";
    assert_eq!(parse(src).found.iter().map(|f| f.0).collect::<Vec<_>>(), vec![0xdebf]);
}
