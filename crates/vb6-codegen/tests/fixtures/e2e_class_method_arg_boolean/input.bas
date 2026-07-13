Attribute VB_Name = "Module1"
Sub Main()
    Dim o As New Class1
    Dim bo As Boolean
    o.TakeBoolByRef bo
    o.TakeBoolByVal bo
End Sub
