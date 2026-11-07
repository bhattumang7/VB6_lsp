use crate::session::{Session, TextEdit};

/// Apply span edits to a source string (right-to-left so offsets stay valid).
fn apply(src: &str, edits: &[TextEdit]) -> String {
    let mut bytes = src.as_bytes().to_vec();
    let mut es: Vec<&TextEdit> = edits.iter().collect();
    es.sort_by_key(|e| std::cmp::Reverse(e.span.start));
    for e in es {
        let s = e.span.start as usize;
        let end = s + e.span.len as usize;
        bytes.splice(s..end, e.new_text.bytes());
    }
    String::from_utf8(bytes).unwrap()
}

fn formatted(src: &str) -> String {
    let s = Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())]);
    apply(src, &s.format(0))
}

#[test]
fn keyword_case_and_indent() {
    let out = formatted("sub foo()\nx = 1\nend sub\n");
    assert_eq!(out, "Sub foo()\n    x = 1\nEnd Sub\n");
}

#[test]
fn select_case_two_levels() {
    let src = "Select Case x\nCase 1\nFoo\nCase Else\nBar\nEnd Select\n";
    let out = formatted(src);
    assert_eq!(
        out,
        "Select Case x\n    Case 1\n        Foo\n    Case Else\n        Bar\nEnd Select\n"
    );
}

#[test]
fn nested_if_inside_sub() {
    let src = "Sub S()\nIf a Then\nb = 1\nElse\nb = 2\nEnd If\nEnd Sub\n";
    let out = formatted(src);
    assert_eq!(
        out,
        "Sub S()\n    If a Then\n        b = 1\n    Else\n        b = 2\n    End If\nEnd Sub\n"
    );
}

#[test]
fn trims_trailing_and_blank_lines() {
    let out = formatted("Sub Foo()   \n\n   \nEnd Sub\n");
    assert_eq!(out, "Sub Foo()\n\n\nEnd Sub\n");
}

#[test]
fn keyword_in_string_is_left_alone() {
    // "end sub" lives in a string literal and is not a keyword token.
    let src = "x = \"end sub\"\n";
    assert_eq!(formatted(src), src);
}

#[test]
fn member_named_like_keyword_not_recased() {
    // `.name` is a member access, not the `Name` keyword.
    let src = "x = obj.name\n";
    assert_eq!(formatted(src), src);
}
