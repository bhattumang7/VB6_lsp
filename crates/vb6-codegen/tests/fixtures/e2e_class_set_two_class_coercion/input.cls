VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Class1"
Attribute VB_Exposed = False
Public Function GetOther() As Object
    Set GetOther = New Class2
End Function
