VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_Exposed = False
Private mObj As Object
Public Property Set PG(ByVal v As Object)
    Set mObj = v
End Property
