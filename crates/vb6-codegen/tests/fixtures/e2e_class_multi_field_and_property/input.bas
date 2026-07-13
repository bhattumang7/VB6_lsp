Attribute VB_Name = "Module1"
Sub Main()
    Dim o As New Class1
    Dim x As Long
    o.F1 = 1
    x = o.F1
    o.F5 = 5
    x = o.F5
    o.G2 = 2
    x = o.G2
End Sub
