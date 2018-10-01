Sub Macro1()
'
' Macro1 Macro
' This Macro is intended for the MAXIENT REPORT. IT clears the heading rows as well as consolidates Duplicate records. 
' Created By Nick Stone 3/14/18 Happy Pi Day!



    Rows("1:8").Select
    Range("A8").Activate
    Selection.Delete Shift:=xlUp
    Columns("A:H").Select
    ActiveSheet.Range("$A$1:$H$1000").RemoveDuplicates Columns:=1, Header:=xlYes
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet 1!R1C1:R1048576C8", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Student ID"), "Count of Student ID", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Student ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    Columns("A:B").Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$B$100").AutoFilter Field:=1, Criteria1:=Array( _
        "000601447", "000603040", "000608765", "000608923", "000614172", "000638659", _
        "000638868", "000645012", "000645259", "000645288", "000645464", "000646513", _
        "000647652", "000648597", "000648832", "000651718", "000652389", "000653478", _
        "000654091"), Operator:=xlFilterValues
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D8").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Sheets("Sheet3").Select
    Application.CutCopyMode = False
    Sheets("Sheet3").Move Before:=Sheets(1)
End Sub


