//! The collection of `For Each` and the variables after `Next` are target
//! expressions: a name or a chain of members, indexes and calls, never a
//! group, a literal, an operator expression or `New`.

mod common;
use common::{offset_of, parse, Parsed};
use vb6_syntax::frontend::ast::ExprNode;

fn for_each(collection: &str) -> (Parsed, String) {
    let src = format!("Sub S()\nFor Each x In {collection}\nNext\nEnd Sub\n");
    (parse(&src), src)
}

#[test]
fn collection_may_be_any_target_expression() {
    for text in ["a", "a.b", "a!b", "a.b.c(1).d", "Foo()", "a(1)"] {
        let (p, _) = for_each(text);
        assert_eq!(p.found, vec![], "{text}");
        let ExprNode::ForEach { collection, .. } = p.node(p.proc_body(0)[0]) else { panic!("expected ForEach") };
        assert!(
            matches!(p.node(*collection), ExprNode::NameRef { .. } | ExprNode::MemberAccess { .. } | ExprNode::Call { .. }),
            "{text}"
        );
    }
}

#[test]
fn collection_that_is_not_a_target_expression_is_a_syntax_error() {
    for text in ["(a)", "a + b", "1", "\"s\"", "New Foo", "Nothing", "UBound(a)", "a, b"] {
        let (p, src) = for_each(text);
        assert!(!p.found.is_empty(), "{text}");
        // The first report is at the offending operand (or just after the part that parsed).
        assert!(p.found[0].1 >= offset_of(&src, " In ", 0) + 4, "{text}: {:?}", p.found);
        assert!(matches!(p.found[0].0, 0x9c70 | 0x9c6f), "{text}: {:?}", p.found);
    }
}

#[test]
fn next_takes_target_expressions() {
    for text in ["a.b", "i.b", "i(1)"] {
        let src = format!("Sub S()\nFor i = 1 To 2\nNext {text}\nEnd Sub\n");
        let p = parse(&src);
        // Parsed, but not the control variable of the loop.
        assert_eq!(p.found.len(), 1, "{text}: {:?}", p.found);
        assert_eq!(p.found[0].0, 0x9ccd, "{text}");
        assert_eq!(p.found[0].1, offset_of(&src, text, 0), "{text}");
    }
}

#[test]
fn next_with_a_value_that_is_not_a_target_expression_is_a_syntax_error() {
    for text in ["1", "i + 1", "(i)"] {
        let src = format!("Sub S()\nFor i = 1 To 2\nNext {text}\nEnd Sub\n");
        assert!(!parse(&src).found.is_empty(), "{text}");
    }
}

#[test]
fn nested_loops_closed_by_one_next_keep_their_order() {
    let ok = parse("Sub S()\nFor i = 1 To 2\nFor j = 1 To 2\nNext j, i\nEnd Sub\n");
    assert_eq!(ok.found, vec![]);
    let wrong = "Sub S()\nFor i = 1 To 2\nFor j = 1 To 2\nNext i, j\nEnd Sub\n";
    let p = parse(wrong);
    assert!(!p.found.is_empty());
    assert!(p.found.iter().all(|d| d.0 == 0x9ccd), "{:?}", p.found);
}

#[test]
fn array_with_brackets_is_a_collection_but_bare_array_and_other_built_ins_are_not() {
    let (p, _) = for_each("Array(1)");
    assert_eq!(p.found, vec![]);
    for text in ["Array", "UBound(a)", "Date", "DoEvents", "CDec(1)"] {
        let (p, _) = for_each(text);
        assert!(!p.found.is_empty(), "{text}");
    }
}

#[test]
fn a_colon_ends_the_next_list() {
    let p = parse("Sub S()\nFor i = 1 To 2\nNext: y = 1\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    assert_eq!(p.proc_body(0).len(), 2);
}

#[test]
fn next_with_no_variable_does_not_end_a_part_of_a_single_line_if() {
    let src = "Sub S()\nIf x Then For i = 1 To 2: Next Else y = 1\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found[0], (0x9c70, offset_of(src, "Else", 0), 4));
    let ok = parse("Sub S()\nIf x Then For i = 1 To 2: Next i Else y = 1\nEnd Sub\n");
    assert_eq!(ok.found, vec![]);
}
