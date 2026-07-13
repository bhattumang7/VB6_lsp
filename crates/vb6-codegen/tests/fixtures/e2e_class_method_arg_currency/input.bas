Attribute VB_Name = "Module1"
Sub Main()
    Dim o As New Class1
    Dim c As Currency
    o.TakeCurrencyByRef c
    o.TakeCurrencyByVal c
End Sub
