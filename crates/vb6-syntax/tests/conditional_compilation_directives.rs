//! Structure errors in `#If` / `#ElseIf` / `#Else` / `#End If`.

mod common;
use common::{offset_of, parse};
use vb6_syntax::frontend::ast::ExprNode;

const WITHOUT_IF: u32 = 0xdebd;
const DUPLICATE_ELSE: u32 = 0xdebe;
const ELSEIF_AFTER_ELSE: u32 = 0xdebf;
const NOT_TERMINATED: u32 = 0xdec0;

#[test]
fn directives_with_no_open_if_are_reported() {
    // A lone `#Else` is reported as a repeated one and a lone `#ElseIf` as one after an `#Else`.
    for (line, id, len) in [("#End If", WITHOUT_IF, 4), ("#Else", DUPLICATE_ELSE, 5), ("#ElseIf 1 Then", ELSEIF_AFTER_ELSE, 7)] {
        let src = format!("{line}\nSub S()\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![(id, 0, len)], "for {line}");
        assert_eq!(p.top.len(), 1, "the procedure after `{line}` is still parsed");
    }
}

#[test]
fn second_else_and_elseif_after_else_are_reported() {
    let src = "#If 1 Then\n#Else\n#Else\n#End If\n";
    assert_eq!(parse(src).found, vec![(DUPLICATE_ELSE, offset_of(src, "#Else", 1), 5)]);

    let src = "#If 1 Then\n#Else\n#ElseIf 1 Then\n#End If\n";
    assert_eq!(parse(src).found, vec![(ELSEIF_AFTER_ELSE, offset_of(src, "#ElseIf", 0), 7)]);
}

#[test]
fn an_if_left_open_at_the_end_is_reported_once_on_the_last_line() {
    let src = "#If 1 Then\n#If 0 Then\nSub S()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(NOT_TERMINATED, offset_of(src, "End Sub", 0), 7)]);
    assert_eq!(p.top.len(), 0, "the procedure is inside the inactive inner branch");

    let src = "#If 1 Then\nSub S()\nEnd Sub\n";
    let p = parse(src);
    assert_eq!(p.found, vec![(NOT_TERMINATED, offset_of(src, "End Sub", 0), 7)]);
    assert_eq!(p.top.len(), 1);
}

#[test]
fn well_formed_blocks_are_clean() {
    let p = parse(
        "#If 0 Then\nSub A()\nEnd Sub\n#ElseIf 1 Then\nSub B()\nEnd Sub\n#Else\nSub C()\nEnd Sub\n#End If\n#If 1 Then\n#If 0 Then\n#End If\n#End If\n",
    );
    assert_eq!(p.found, vec![]);
    assert_eq!(p.top.len(), 1);
    let ExprNode::ProcDecl { name, .. } = p.node(p.top[0]) else { panic!("expected procedure") };
    assert_eq!(p.name(*name), "B");
}
