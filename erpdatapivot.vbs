'
' Macro2 Macro
'

'
    Selection.CurrentRegion.Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet2!R1C1:R791C13", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:="Sheet3!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion10
    Sheets("Sheet3").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Fiscal Year")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Fiscal Quarter")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sales Rep Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Food Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Net Booking"), "Sum of Net Booking", xlSum
End Sub

