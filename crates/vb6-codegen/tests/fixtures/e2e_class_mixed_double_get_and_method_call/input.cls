VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Class1"
Attribute VB_Exposed = False
Private mP As Double
Public Property Get P() As Double
    P = mP
End Property
Public Property Let P(v As Double)
    mP = v
End Property
Public Sub DoIt(x As Long)
    Dim z As Long
    z = x
End Sub
