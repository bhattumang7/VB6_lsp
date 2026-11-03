//! Valid VB6 that once drew parse or binding errors: every source here must
//! analyse without a single diagnostic.

use vb6_engine::session::Session;

fn diagnostics(files: &[(&str, &str)]) -> Vec<Vec<u32>> {
    let s = Session::from_sources(
        files.iter().map(|(n, c)| (n.to_string(), c.as_bytes().to_vec())).collect(),
    );
    (0..files.len()).map(|i| s.diagnostics(i).iter().map(|d| d.code).collect()).collect()
}

fn assert_clean(name: &str, src: &str) {
    assert_eq!(diagnostics(&[(name, src)]), vec![Vec::<u32>::new()], "{name}");
}

#[test]
fn bare_end_inside_block_if() {
    assert_clean(
        "m.bas",
        "Option Explicit\nPublic Sub QuitIfAsked(ByVal bQuit As Boolean)\n    If bQuit Then\n        End ' Terminate program\n    End If\nEnd Sub\n",
    );
}

#[test]
fn static_procedure_and_its_call() {
    assert_clean(
        "m.bas",
        "Option Explicit\nPrivate Static Sub CountCalls()\n    Dim n As Integer\n    n = n + 1\nEnd Sub\n\nPublic Sub UseCountCalls()\n    CountCalls\nEnd Sub\n",
    );
}

#[test]
fn local_variable_named_error() {
    assert_clean(
        "m.bas",
        "Option Explicit\nPublic Function GetCode() As Long\n    Dim Error As Long\n    Error = 5\n    If Error <> 0 Then GetCode = Error\nEnd Function\n",
    );
}

#[test]
fn redim_member_array_in_with() {
    assert_clean(
        "m.bas",
        "Option Explicit\nPrivate Type Holder\n    Items() As String\nEnd Type\n\nPrivate mHolder As Holder\n\nPublic Sub Reset(ByVal n As Long)\n    With mHolder\n        ReDim .Items(n)\n        ReDim Preserve .Items(n)\n    End With\nEnd Sub\n",
    );
}

#[test]
fn undefined_with_member_in_redim_is_still_reported() {
    let codes = diagnostics(&[(
        "m.bas",
        "Option Explicit\nPrivate Type Holder\n    Items() As String\nEnd Type\nPrivate mHolder As Holder\nPublic Sub Reset()\n    With mHolder\n        ReDim .Nope(3)\n    End With\nEnd Sub\n",
    )]);
    assert_eq!(codes, vec![vec![461]]);
}

#[test]
fn directive_wrapping_a_case_clause() {
    assert_clean(
        "m.bas",
        "Option Explicit\nPublic Function Factor(ByVal s As String) As Integer\n    Select Case s\n        Case \"u\": Factor = -6\n        #If UNDEFINED = 0 Then\n        Case \"U\": Factor = -6\n        #End If\n        Case \"m\": Factor = -3\n    End Select\nEnd Function\n",
    );
}

#[test]
fn print_method_semicolons() {
    assert_clean(
        "m.bas",
        "Option Explicit\nPublic Sub PrintTo(ByVal OutObj As Object, ByVal Frm As Form, ByVal word As String)\n    Printer.Print \" \" + word$;\n    Printer.Print \"A\"; word\n    Frm.Print \" \" + word$;\n    Frm.Print \"A\"; word\n    OutObj.Print \" \" + word$;\n    OutObj.Print \"A\"; word\n    OutObj.Print (\" \" + word$);\nEnd Sub\n",
    );
}

#[test]
fn graphics_methods_with_step_and_box_flags() {
    assert_clean(
        "m.bas",
        "Option Explicit\nPublic Sub Draw(ByVal Frm As Form, ByVal OutObj As Object)\n    Frm.Line (0, 0)-(100, 100), , B\n    Frm.Line (0, 0)-(100, 100), vbRed, BF\n    OutObj.Line (0, 0)-(100, 100), , B\n    Frm.Line Step(1, 1)-Step(10, 10)\n    Frm.Line -Step(10, 10)\n    Frm.Circle Step(5, 5), 10\n    Frm.PSet Step(1, 1)\nEnd Sub\n",
    );
}

#[test]
fn form_module_graphics_statements() {
    assert_clean(
        "Form1.frm",
        "VERSION 5.00\nBegin VB.Form Form1\n   Caption = \"x\"\nEnd\nAttribute VB_Name = \"Form1\"\nOption Explicit\nPrivate Sub Draw(x As Long)\n Line (0,0)-(1,1), vbRed, B\n Line (0,0)-(1,1), , BF\n Line Step(1,1)-Step(2,2)\n Line -Step(2,2)\n PSet (1, 2), vbRed\n PSet Step(1,2)\n Circle (1,2), 3\n Circle Step(1,2), 3\n Me.Line (0,0)-Step(5,5), , B\n Print \"a\"; x;\n Me.Print \"a\"; x;\nEnd Sub\n",
    );
}
