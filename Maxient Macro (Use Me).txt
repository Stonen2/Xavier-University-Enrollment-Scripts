Sub Macro1()
'
' Macro1 Macro
' This Macro is intended for the MAXIENT REPORT. IT clears the heading rows as well as consolidates Duplicate records. 
' Created By Nick Stone 3/14/18 Happy Pi Day!



    Rows("1:8").Select
    Range("A8").Activate
    Selection.ClearContents
    Selection.Delete Shift:=xlUp
    Columns("A:H").Select
    ActiveSheet.Range("$A$1:$H$28").RemoveDuplicates Columns:=1, Header:=xlYes
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
    Range("A3:B22").Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D6").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Sheets("Sheet2").Select
    Application.CutCopyMode = False
    Sheets("Sheet2").Move Before:=Sheets(1)
    ChDir "C:\Users\stonen2\AppData\Local\Temp"
End Sub


