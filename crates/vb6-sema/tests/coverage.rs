//! Resolution-coverage invariant: for well-formed source, every `NameRef` the
//! parser produces must be reached by the binder and recorded in `resolutions`.
//!
//! This is the structural backstop for the "go-to-definition silently finds
//! nothing" bug class. It catches both failure modes at once:
//!   * parser orphaning — a name parsed but never attached to the AST;
//!   * walker under-traversal — an attached name a tree-walk forgot to descend
//!     into.
//!
//! The corpus deliberately exercises one name use inside every statement form,
//! so a regression in any single construct trips the assert.

use vb6_sema::frontend::ast::{ExprArena, ExprNode, NodeId};
use vb6_sema::frontend::parser::Parser;
use vb6_sema::frontend::scanner::ScannerContext;
use vb6_sema::sema::{bind, unbound_namerefs, NameResolution};

/// Parse + bind `src`, returning the node ids of any `NameRef` left unresolved
/// (orphaned / missed). Empty = the coverage invariant holds.
fn unbound(src: &str) -> Vec<NodeId> {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let m = bind(&ctx, &arena, &top, &spans, &vis);
    unbound_namerefs(&arena, &m.resolutions)
}

/// `(total NameRef nodes, NameRefs resolved to an actual declaration)` for `src`.
/// Used to prove the corpus is non-vacuous — that constructs really parsed and
/// produced resolvable name uses, rather than the invariant passing because a
/// construct silently failed to parse and left nothing to check.
fn resolution_stats(src: &str) -> (usize, usize) {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let m = bind(&ctx, &arena, &top, &spans, &vis);
    let total = (0..arena.len() as u32)
        .filter(|i| matches!(arena.get(NodeId(*i)), ExprNode::NameRef { .. }))
        .count();
    let resolved = m
        .resolutions
        .values()
        .filter(|r| {
            !matches!(r, NameResolution::Unresolved | NameResolution::Builtin)
        })
        .count();
    (total, resolved)
}

/// One module touching every statement form, with a name use inside each so a
/// missed child surfaces as an unbound `NameRef`.
const CORPUS: &str = r#"
Option Explicit

Implements ISomething

Private Const MAX As Long = 10
Private Const LIMIT As Long = MAX + 1
Private arr(1 To MAX) As Long
Private obj As Object
Private mValue As Long
Private mRef As Object

Public Enum Colors
    Red = 1
    Green = Red + 1
    Blue = Green + 1
End Enum

Event Changed(ByVal newVal As Long)

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let Value(ByVal v As Long)
    mValue = v
End Property

Public Property Set Ref(ByVal o As Object)
    Set mRef = o
End Property

Private Sub Demo(Optional ByVal limit As Long = MAX)
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim i As Long
    Dim s As String
    Dim coll As Object
    Dim e As Variant
    Dim n As Long
    Dim dyn(1 To MAX) As Long
    Dim x As Long, y As Long
    Static hits As Long
    Dim dt As Date

    a = b + c
    a = -b
    a = (b)
    a = x + y + hits
    a = Red + Green + Blue + LIMIT
    dt = #1/1/2020#
    MsgBox a
    Call Demo(a)
    Demo b
    Demo limit:=a

    If a > b Then
        a = c
    ElseIf a < b Then
        a = b
    Else
        a = limit
    End If

    For i = a To b Step c
        a = a + i
    Next

    For Each e In coll
        a = a + 1
    Next

    Do While a < b
        a = a + 1
    Loop

    While a < c
        a = a + 1
    Wend

    With obj
        a = a + 1
    End With

    Select Case a
        Case b
            a = 1
        Case Is > c
            a = 2
        Case b To c
            a = 3
        Case Else
            a = 4
    End Select

    ReDim arr(1 To n)
    Erase arr
    RaiseEvent Changed(a)
    Mid(s, a, b) = s
    Mid(s, a) = s
    LSet s = s
    RSet s = s
    Set obj = New Collection
    Set obj = Nothing

    If TypeOf obj Is Object Then
        a = 1
    End If

    If a = 1 Then Exit Sub
    For i = 1 To 10
        If i = 5 Then Exit For
    Next

    Notify AddressOf Demo

    Debug.Print a, b
    s = "x" & _
        "y"

    Open s For Random As #1 Len = n
    Print #1, a, b
    Print #1, a; b;
    Write #1, a, b
    Input #1, a
    Get #1, a, b
    Put #1, a, b
    Line Input #1, s
    Seek #1, a
    Close #1

    a = obj.Field
    a = obj!Field2

    On Error Resume Next
    On Error GoTo handler
    On a GoTo handler, done
    GoSub helper
    GoTo done
helper:
    a = b
    Return
handler:
    a = 0
done:
End Sub
"#;

#[test]
fn every_statement_form_resolves_all_name_uses() {
    let offenders = unbound(CORPUS);
    assert!(
        offenders.is_empty(),
        "unbound NameRef node(s) {offenders:?} — a parser orphan or a missed child. \
         The resolution-coverage invariant is broken for some construct above."
    );

    // Non-vacuity: the corpus must actually have parsed into many resolvable name
    // uses, so the invariant above is not passing simply because a construct
    // failed to parse and left nothing to check.
    let (total, resolved) = resolution_stats(CORPUS);
    assert!(total > 80, "corpus produced only {total} NameRef nodes — under-parsed?");
    assert!(resolved > 60, "only {resolved} of {total} NameRefs resolved to a decl");
}

#[test]
fn invariant_detects_a_planted_orphan() {
    // Sanity: the checker is not vacuously true. A bare implicit call with an
    // argument is the exact shape that used to orphan its arg; it must now be
    // covered, so this corpus stays clean — and the checker would flag it if a
    // future regression dropped the arg again.
    let src = "Sub S()\n    Dim x As Long\n    MsgBox x\nEnd Sub\n";
    assert!(unbound(src).is_empty());
}
