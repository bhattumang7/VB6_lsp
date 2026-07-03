Attribute VB_Name = "Module1"
Sub Main()
    Call Foo(F())
End Sub
Function F() As Long
    F = 7
End Function
Sub Foo(ByVal x As Long)
    Dim z As Long
    z = x
End Sub
