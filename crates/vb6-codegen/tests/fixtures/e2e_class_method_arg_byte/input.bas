Attribute VB_Name = "Module1"
Sub Main()
    Dim o As New Class1
    Dim by As Byte
    o.TakeByteByRef by
    o.TakeByteByVal by
End Sub
