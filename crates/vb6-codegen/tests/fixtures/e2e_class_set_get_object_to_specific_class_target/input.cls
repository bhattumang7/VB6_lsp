VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Class1"
Attribute VB_Exposed = False
Private mP As Object
Public Property Get P() As Object
    Set P = mP
End Property
