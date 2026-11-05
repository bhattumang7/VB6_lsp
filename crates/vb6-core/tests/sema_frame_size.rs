//! The nonstatic locals of a procedure must fit its frame (32636 bytes): each takes the bytes
//! of its type, and the local that passes the limit is reported.

use vb6_core::session::Session;

const TOO_MANY_LOCALS: u32 = 57028;

fn codes(header: &str, footer: &str, decl: &dyn Fn(usize) -> String, n: usize) -> Vec<(u32, u32, u32)> {
    let mut src = format!("{header}\n");
    for i in 0..n {
        src.push_str(&decl(i));
        src.push('\n');
    }
    src.push_str(&format!("{footer}\nSub Main()\nEnd Sub\n"));
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    s.diagnostics(0).iter().map(|d| (d.code, d.span.start, d.span.len)).collect()
}

fn sub(ty: &str) -> impl Fn(usize) -> String + '_ {
    move |i| format!("Dim v{i}{ty}")
}

#[test]
fn locals_fill_the_frame_by_the_bytes_of_their_type() {
    // (type text, size in bytes, most that fit)
    for (ty, fits) in [
        (" As Long", 8159),
        (" As Single", 8159),
        (" As String", 8159),
        (" As Object", 8159),
        (" As Integer", 16318),
        (" As Boolean", 16318),
        (" As Byte", 16318),
        (" As Double", 4079),
        (" As Currency", 4079),
        (" As Date", 4079),
        (" As Variant", 2039),
        (" () As Long", 8159),
        (" As String * 10", 1631),
        (" As String * 11", 1483),
        ("(3) As Long", 1019),
        ("(1 To 2, 1 To 2) As Double", 815),
    ] {
        let d = sub(ty);
        let found = codes("Sub Foo()", "End Sub", &d, fits);
        assert!(!found.iter().any(|f| f.0 == TOO_MANY_LOCALS), "{ty} x {fits}");
        let found = codes("Sub Foo()", "End Sub", &d, fits + 1);
        let hit: Vec<_> = found.iter().filter(|f| f.0 == TOO_MANY_LOCALS).collect();
        assert_eq!(hit.len(), 1, "{ty} x {}", fits + 1);
    }
}

#[test]
fn the_function_value_takes_its_place_in_the_frame_and_static_locals_and_parameters_do_not() {
    let long = sub(" As Long");
    // A Double result takes 8 bytes: two Long locals fewer fit.
    let f = |n| codes("Function Foo() As Double", "End Function", &long, n);
    assert!(!f(8157).iter().any(|c| c.0 == TOO_MANY_LOCALS));
    assert!(f(8158).iter().any(|c| c.0 == TOO_MANY_LOCALS));
    // A Variant result takes 16.
    let g = |n| codes("Function Foo() As Variant", "End Function", &long, n);
    assert!(!g(8155).iter().any(|c| c.0 == TOO_MANY_LOCALS));
    assert!(g(8156).iter().any(|c| c.0 == TOO_MANY_LOCALS));
    // Parameters are not part of the frame.
    let h = |n| codes("Sub Foo(a As Long, b As Long, c As Double)", "End Sub", &long, n);
    assert!(!h(8159).iter().any(|c| c.0 == TOO_MANY_LOCALS));
    assert!(h(8160).iter().any(|c| c.0 == TOO_MANY_LOCALS));
    // Static locals take no frame space.
    let stat = |i| format!("Static v{i} As Long");
    assert!(!codes("Sub Foo()", "End Sub", &stat, 9000).iter().any(|c| c.0 == TOO_MANY_LOCALS));
}

#[test]
fn the_local_that_passes_the_limit_is_the_one_reported() {
    let long = sub(" As Long");
    let found = codes("Sub Foo()", "End Sub", &long, 8161);
    let hit: Vec<_> = found.iter().filter(|f| f.0 == TOO_MANY_LOCALS).collect();
    assert_eq!(hit.len(), 1);
    let mut src = String::from("Sub Foo()\n");
    for i in 0..8161 {
        src.push_str(&format!("Dim v{i} As Long\n"));
    }
    let at = src.find("v8159 ").unwrap() as u32;
    assert_eq!(hit[0].1, at);
}
