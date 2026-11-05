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

#[test]
fn loops_with_and_select_case_take_hidden_frame_space() {
    let long = sub(" As Long");
    // (the lines before the locals, what closes them, most Long locals that fit)
    let cases: [(&str, &str, usize); 16] = [
        ("Sub Foo()\nDim i As Long\nFor i = 1 To 2\nNext i", "End Sub", 8156),
        ("Sub Foo()\nDim i As Integer\nFor i = 1 To 2\nNext i", "End Sub", 8157),
        ("Sub Foo()\nDim i As Double\nFor i = 1 To 2\nNext i", "End Sub", 8153),
        ("Sub Foo()\nDim i As Single\nFor i = 1 To 2\nNext i", "End Sub", 8156),
        ("Sub Foo()\nDim i As Currency\nFor i = 1 To 2\nNext i", "End Sub", 8153),
        ("Sub Foo()\nDim v As Variant, c As Object\nFor Each v In c\nNext v", "End Sub", 8153),
        ("Sub Foo()\nDim v As Object, c As Object\nFor Each v In c\nNext v", "End Sub", 8156),
        ("Sub Foo()\nDim o As Object\nWith o\nEnd With", "End Sub", 8157),
        ("Sub Foo()\nDim i As Long\nSelect Case i\nCase 1\nEnd Select", "End Sub", 8157),
        ("Sub Foo()\nDim i As String\nSelect Case i\nCase \"a\"\nEnd Select", "End Sub", 8157),
        ("Sub Foo()\nDim i As Integer\nSelect Case i\nCase 1\nEnd Select", "End Sub", 8158),
        ("Sub Foo()\nDim i As Double\nSelect Case i\nCase 1\nEnd Select", "End Sub", 8155),
        ("Sub Foo()\nDim i As Currency\nSelect Case i\nCase 1\nEnd Select", "End Sub", 8155),
        ("Sub Foo()\nDim i As Long\nFor i = 1 To 2\nNext i\nFor i = 1 To 2\nNext i", "End Sub", 8154),
        ("Sub Foo()\nDim i As Long, j As Long\nFor i = 1 To 2\nFor j = 1 To 2\nNext j\nNext i", "End Sub", 8153),
        ("Sub Foo()\nDim o As Object\nWith o\nEnd With\nWith o\nEnd With", "End Sub", 8156),
    ];
    for (header, footer, fits) in cases {
        // The locals come first, then the statements, as the measurement placed them.
        let (decls, stmts) = match header.find("\nFor ").or_else(|| header.find("\nWith ")).or_else(|| header.find("\nSelect ")) {
            Some(_) => {
                let lines: Vec<&str> = header.split('\n').collect();
                let first_stmt = lines.iter().position(|l| l.starts_with("For ") || l.starts_with("With ") || l.starts_with("Select ")).unwrap();
                (lines[..first_stmt].join("\n"), lines[first_stmt..].join("\n"))
            }
            None => (header.to_string(), String::new()),
        };
        let run = |n: usize| {
            let mut body = String::new();
            for i in 0..n {
                body.push_str(&format!("{}\n", long(i)));
            }
            let src = format!("{decls}\n{body}{stmts}\n{footer}\nSub Main()\nEnd Sub\n");
            let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
            s.diagnostics(0).iter().any(|d| d.code == TOO_MANY_LOCALS)
        };
        assert!(!run(fits), "{header:?} x {fits}");
        assert!(run(fits + 1), "{header:?} x {}", fits + 1);
    }
}

fn param_array_calls(statements: usize, args: usize) -> bool {
    // Lines of at most 100 arguments: a longer physical line is cut.
    let lines: Vec<String> = vec!["1"; args].chunks(100).map(|c| c.join(", ")).collect();
    let call = format!("Foo {}", lines.join(", _\n"));
    let body = vec![call; statements].join("\n");
    let src = format!("Sub Foo(ParamArray a())\nEnd Sub\nSub Main()\n{body}\nEnd Sub\n");
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    s.diagnostics(0).iter().any(|d| d.code == TOO_MANY_LOCALS)
}

#[test]
fn param_array_arguments_take_sixteen_bytes_each_and_later_calls_reuse_them() {
    assert!(!param_array_calls(1, 2039));
    assert!(param_array_calls(1, 2040));
    // Two calls share the same temporaries.
    assert!(!param_array_calls(2, 1500));
    assert!(!param_array_calls(2, 2039));
    assert!(param_array_calls(2, 2040));
}

fn variant_chain(terms: usize) -> bool {
    let lines: Vec<String> = vec!["v"; terms].chunks(100).map(|c| c.join(" + ")).collect();
    let src = format!("Sub Main()\nDim x As Long, v As Variant\nx = {}\nEnd Sub\n", lines.join(" + _\n"));
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    s.diagnostics(0).iter().any(|d| d.code == TOO_MANY_LOCALS)
}

#[test]
fn every_operation_on_a_variant_takes_a_sixteen_byte_temporary() {
    assert!(!variant_chain(2039));
    assert!(variant_chain(2040));
}

#[test]
fn the_longest_statement_the_compiler_accepts_is_analyzed_without_overflowing_the_stack() {
    let lines: Vec<String> = vec!["1"; 2552].chunks(250).map(|c| c.join(" + ")).collect();
    let src = format!("Sub Main()\nDim x As Long\nx = {}\nEnd Sub\n", lines.join(" + _\n"));
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    assert_eq!(s.diagnostics(0).len(), 0);
}

#[test]
fn deeply_nested_brackets_are_analyzed_without_overflowing_the_stack() {
    let src = format!("Sub Main()\nDim x As Long\nx = {}1{}\nEnd Sub\n", "(".repeat(480), ")".repeat(480));
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    assert_eq!(s.diagnostics(0).len(), 0);
}

#[test]
fn brackets_nested_four_thousand_deep_are_analyzed() {
    // Lines of 400 brackets, joined by continuations.
    let opens: Vec<String> = (0..10).map(|_| "(".repeat(400)).collect();
    let closes: Vec<String> = (0..10).map(|_| ")".repeat(400)).collect();
    let src = format!(
        "Sub Main()\nDim x As Long\nx = {} 1 _\n{}\nEnd Sub\n",
        opens.join(" _\n"),
        closes.join(" _\n")
    );
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    assert_eq!(s.diagnostics(0).len(), 0);
}

fn boolean_concat(term: &str, n: usize) -> bool {
    let lines: Vec<String> = vec![term; n].chunks(60).map(|c| c.join(" & ")).collect();
    let src = format!(
        "Sub Main()\nDim x As String, y As Long, s As String\nx = {}\nEnd Sub\n",
        lines.join(" & _\n")
    );
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    s.diagnostics(0).iter().any(|d| d.code == TOO_MANY_LOCALS)
}

#[test]
fn a_boolean_joined_with_ampersand_takes_a_thirty_two_byte_temporary() {
    for term in ["(y = y)", "(y < y)", "(s = s)", "(1 = 2)", "(True)"] {
        assert!(!boolean_concat(term, 1020), "{term}");
        assert!(boolean_concat(term, 1021), "{term}");
    }
    // Operands that are not Boolean are not converted that way.
    assert!(!boolean_concat("y", 2000));
}

fn variant_param_calls(param: &str, arg: &str, n: usize) -> bool {
    let lines: Vec<String> = vec![format!("g({arg})"); n].chunks(80).map(|c| c.join(" + ")).collect();
    let src = format!(
        "Function g({param}) As Long\nEnd Function\nSub Main()\nDim x As Long, y As Long\nx = {}\nEnd Sub\n",
        lines.join(" + _\n")
    );
    let s = Session::from_sources(vec![("m.bas".to_string(), src.into_bytes())]);
    s.diagnostics(0).iter().any(|d| d.code == TOO_MANY_LOCALS)
}

#[test]
fn an_integer_literal_passed_by_reference_to_a_variant_takes_a_thirty_two_byte_temporary() {
    for param in ["Optional v As Variant", "v As Variant"] {
        assert!(!variant_param_calls(param, "1", 1019), "{param}");
        assert!(variant_param_calls(param, "1", 1020), "{param}");
    }
    // By value, or from a variable, no temporary is needed.
    assert!(!variant_param_calls("ByVal v As Variant", "1", 1500));
    assert!(!variant_param_calls("v As Variant", "y", 1500));
}
