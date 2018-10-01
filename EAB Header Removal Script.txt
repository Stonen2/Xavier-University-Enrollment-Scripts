Sub EABReport()
'
' EABReport Macro
' This macro is intended for EAB reports. It clears the heading Rows as well as consolidates duplicate records.  Created by Nick Stone 3/14/18  Happy Pi Day!
'

'
    Rows("1:8").Select
    Range("A8").Activate
    Selection.Delete Shift:=xlUp
    Columns("A:H").Select
    ActiveSheet.Range("$A$1:$H$1000").RemoveDuplicates Columns:=1, Header:=xlYes
End Sub
