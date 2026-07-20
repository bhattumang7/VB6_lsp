Attribute VB_Name = "Module1"
Sub Main()
    Dim o As Class1
    Dim other As New Class1
    Set o = New Class1
    Set o = other
    Set o = Nothing
End Sub
