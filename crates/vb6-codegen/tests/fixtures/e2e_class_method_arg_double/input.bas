Attribute VB_Name = "Module1"
Sub Main()
    Dim o As New Class1
    Dim d As Double
    o.TakeDoubleByRef d
    o.TakeDoubleByVal d
End Sub
