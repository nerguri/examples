Sub Macro1()
'
' Macro1 Macro
'
'
    Range("B11:K11").Select
    Selection.AutoFill Destination:=Range("B11:K16"), Type:=xlFillDefault
    Range("B11:K16").Select
    Range("K2:K16").Select
    Selection.AutoFill Destination:=Range("K2:P16"), Type:=xlFillDefault
    Range("K2:P16").Select
    Columns("B:P").Select
    Selection.ColumnWidth = 4
End Sub

