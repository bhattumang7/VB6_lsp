//! `Call` takes a single bracket group on a name and no `!` member; `RaiseEvent` takes its
//! arguments by position; a `WithEvents` variable is one plain name.

mod common;
use common::*;
use vb6_syntax::frontend::parser::ModuleKind;

const SYNTAX_ERROR: u32 = 0x9c70;
const EXPECTED: u32 = 0x9c6f;

/// The statements of `Sub S()` holding `body`, rendered.
fn statements(body: &str) -> (Parsed, Vec<String>) {
    let src = format!("Sub S()\n{body}\nEnd Sub\n");
    let p = parse(&src);
    let shown = p.proc_body(0).iter().map(|s| p.show(*s)).collect();
    (p, shown)
}

#[test]
fn a_call_target_with_one_bracket_group_is_read_whole() {
    for (text, shown) in [
        ("Call foo(1)", "(Call (call foo [1%]) [])"),
        ("Call foo", "(Call foo [])"),
        ("Call o.a(1)", "(Call (call (dot o a) [1%]) [])"),
        ("Call o.a (1)", "(Call (call (dot o a) [1%]) [])"),
        // A group after a member is the member's, not a second group on the name.
        ("Call o.a(1).bar(2)", "(Call (call (dot (call (dot o a) [1%]) bar) [2%]) [])"),
        ("Call o.a(1).bar", "(Call (dot (call (dot o a) [1%]) bar) [])"),
    ] {
        let (p, stmts) = statements(&format!(" {text}"));
        assert_eq!(p.found, vec![], "{text}");
        assert_eq!(stmts, vec![shown.to_string()], "{text}");
    }
}

#[test]
fn a_call_target_starting_at_a_with_member_takes_further_groups() {
    let (p, stmts) = statements(" With o\n Call .a(1)(2)\n End With");
    assert_eq!(p.found, vec![]);
    assert_eq!(stmts, vec!["(With o {(Call (call (call (dot <with> a) [1%]) [2%]) [])})".to_string()]);
}

#[test]
fn a_second_bracket_group_directly_after_the_first_is_a_syntax_error() {
    for text in ["Call foo(1)(2)", "Call arr(1) (1)", "Call o.a(1)(2)", "Call o(1)(2)(3)", "Call foo(1)(2) : x = 1"] {
        let src = format!("Sub S()\n {text}\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![(SYNTAX_ERROR, offset_of(&src, "Call", 0), 4)], "{text}");
        // The statement is dropped; nothing after it on the line is read.
        assert_eq!(p.proc_body(0).iter().map(|s| p.show(*s)).collect::<Vec<_>>(), vec!["{}".to_string()], "{text}");
    }
    // Inside a single-line If the statement stands in the If.
    let src = "Sub S()\n If True Then Call o.a(1)(2)\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(SYNTAX_ERROR, offset_of(src, "Call", 0), 4)]);
}

#[test]
fn a_bang_member_anywhere_in_a_call_target_is_a_syntax_error_at_the_bang_member() {
    for (text, member) in [
        ("Call o!a", "a"),
        ("Call o!a.b", "a"),
        ("Call o!a(1).b", "a"),
        ("Call o.a!b.x", "b"),
        ("Call o.a!b(1)", "b"),
        ("Call o.a(1).b!c", "c"),
    ] {
        let src = format!("Sub S()\n {text}\nEnd Sub\n");
        let p = parse(&src);
        let at = src.find('!').unwrap() + 1;
        assert_eq!(p.found, vec![(SYNTAX_ERROR, at as u32, member.len() as u32)], "{text}");
    }
    let src = "Sub S()\n With o\n Call .a!b\n End With\nEnd Sub\n";
    assert_eq!(parse(src).found, vec![(SYNTAX_ERROR, src.find('!').unwrap() as u32 + 1, 1)]);
}

#[test]
fn raise_event_arguments_are_positional() {
    let class_body = |text: &str| {
        let src = format!("Event E(a As Long, b As Long)\nSub S()\n RaiseEvent {text}\nEnd Sub\n");
        let p = parse_as(&src, ModuleKind::Class);
        let shown: Vec<String> = p.proc_body(0).iter().map(|s| p.show(*s)).collect();
        (src, p, shown)
    };
    let (_, p, shown) = class_body("E(1, 2)");
    assert_eq!(p.found, vec![]);
    assert_eq!(shown, vec!["(RaiseEvent - E [1% 2%])".to_string()]);
    // A named argument anywhere in the list is a syntax error at the statement, whatever its name.
    for text in ["E(a:=1)", "E(1, b:=2)", "E(z:=1)", "E(a:=1, b:=2)"] {
        let (src, p, shown) = class_body(text);
        assert_eq!(p.found, vec![(SYNTAX_ERROR, offset_of(&src, "E(", 1), 1)], "{text}");
        assert_eq!(shown, vec!["{}".to_string()], "{text}");
    }
    // A named argument to a call inside an argument is the call's own.
    let (_, p, shown) = class_body("E(foo(a:=1), 2)");
    assert_eq!(p.found, vec![]);
    assert_eq!(shown, vec!["(RaiseEvent - E [(call foo [(named a 1%)]) 2%])".to_string()]);
}

#[test]
fn a_with_events_variable_is_a_plain_name() {
    for text in ["Private WithEvents o(3) As Class2", "Private WithEvents o() As Class2"] {
        let src = format!("{text}\n");
        let p = parse_as(&src, ModuleKind::Class);
        let paren = src.find('(').unwrap() as u32;
        assert_eq!(p.found, vec![(EXPECTED, paren, 1)], "{text}");
        assert_eq!(p.messages, vec!["Expected: As".to_string()], "{text}");
        // The declaration is kept without its bounds and type.
        let ExprBlock(items) = block_of(&p, 0);
        assert_eq!(items, vec!["(item o withevents bounds=- type=- init=-)".to_string()], "{text}");
    }
    let p = parse_as("Private WithEvents o As Class2\n", ModuleKind::Class);
    assert_eq!(p.found, vec![]);
    let ExprBlock(items) = block_of(&p, 0);
    assert_eq!(items, vec!["(item o withevents bounds=- type=(usertype Class2 -) init=-)".to_string()]);
}

struct ExprBlock(Vec<String>);

fn block_of(p: &Parsed, index: usize) -> ExprBlock {
    use vb6_syntax::frontend::ast::ExprNode;
    let ExprNode::Block { stmts } = p.node(p.top[index]) else { panic!("not a block") };
    ExprBlock(stmts.iter().map(|s| p.show(*s)).collect())
}
