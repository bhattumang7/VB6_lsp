Attribute VB_Name = "Module1"
Sub Main()
    Dim o As New Class1
    Dim i As Integer
    Dim s As String
    Dim y As Object
    o.TakeInt i
    o.TakeIntByVal i
    o.TakeStr s
    o.TakeObj y
End Sub
