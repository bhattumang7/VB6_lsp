Attribute VB_Name = "Module1"
Sub Main()
    Dim o As New Class1
    Dim d As Date
    o.TakeDateByRef d
    o.TakeDateByVal d
End Sub
