//! Assigning to a constant or an enum member is an error (57082).

use vb6_core::session::Session;

const ASSIGNMENT_TO_CONSTANT: u32 = 0xdefa;

fn session(src: &str) -> Session {
    Session::from_sources(vec![("m.bas".to_string(), src.as_bytes().to_vec())])
}

fn codes(s: &Session) -> Vec<(u32, u32, u32)> {
    s.diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

fn at(src: &str, ctx: &str, member: &str) -> u32 {
    (src.find(ctx).unwrap() + ctx.find(member).unwrap()) as u32
}

#[test]
fn module_and_local_constants_cannot_be_assigned() {
    let src = "Const Limit = 5\nSub Run()\n    Const Local = 1\n    Limit = 6\n    Local = 2\nEnd Sub\n";
    let s = session(src);
    assert_eq!(
        codes(&s),
        vec![
            (ASSIGNMENT_TO_CONSTANT, at(src, "Limit = 6", "Limit"), 5),
            (ASSIGNMENT_TO_CONSTANT, at(src, "Local = 2", "Local"), 5),
        ]
    );
    assert_eq!(s.diagnostics(0)[0].message().as_deref(), Some("Assignment to constant not permitted"));
}

#[test]
fn enum_members_cannot_be_assigned() {
    let src = "Enum Color\n    Red\n    Blue\nEnd Enum\nSub Run()\n    Red = 3\nEnd Sub\n";
    let s = session(src);
    assert_eq!(codes(&s), vec![(ASSIGNMENT_TO_CONSTANT, at(src, "Red = 3", "Red"), 3)]);
}

#[test]
fn variables_and_reads_of_constants_are_fine() {
    let src = "Const Limit = 5\nDim total As Long\nSub Run()\n    Dim n As Long\n    n = Limit\n    total = Limit + 1\n    For n = 1 To Limit\n    Next\nEnd Sub\n";
    let s = session(src);
    assert_eq!(codes(&s), vec![]);
}

#[test]
fn a_constant_as_a_loop_variable_is_assigned() {
    let src = "Const Limit = 5\nSub Run()\n    For Limit = 1 To 3\n    Next\nEnd Sub\n";
    let s = session(src);
    assert_eq!(codes(&s), vec![(ASSIGNMENT_TO_CONSTANT, at(src, "Limit = 1", "Limit"), 5)]);
}
