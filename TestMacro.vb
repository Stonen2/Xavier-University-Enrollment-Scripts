Public Sub TestMacro()
' Created by Nick Stone in complimenet 
' TestMacro Macro
'
'
Dim objExcel As Object
Dim Tester As String
Dim Client As String

Dim Names As String
Set objExcel = CreateObject("Excel.Application")
Dim exWb As Object

Set exWb = objExcel.Workbooks.Open("v.xlsx")

Names = exWb.Worksheets(1).Range("B6")
Tester = exWb.Worksheets(1).Range("B11")
Client = exWb.Worksheets(1).Range("B5")


exWb.Close

Set exWb = Nothing
 
Dim x As String
x = "TEMPORARY TESTING"
Selection.find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.find.ClearFormatting
Selection.find.Replacement.ClearFormatting
With Selection.find
.Text = "#PROJECT NAME#"
.Replacement.Text = Names
.Forward = True
.Wrap = wdFindContinue & vcCrLf
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
'''''''''''''
Selection.find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.find.ClearFormatting
Selection.find.Replacement.ClearFormatting
With Selection.find
.Text = "#PROJECT CLIENT#"
.Replacement.Text = Client
.Forward = True
.Wrap = wdFindContinue & vcCrLf
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
''''''''''

Selection.find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.find.ClearFormatting
Selection.find.Replacement.ClearFormatting
With Selection.find
.Text = "#PROJECT TESTERS#"
.Replacement.Text = Tester
.Forward = True
.Wrap = wdFindContinue & vcCrLf
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
'''''''''''''''
''''''''''

Selection.find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.find.ClearFormatting
Selection.find.Replacement.ClearFormatting
With Selection.find
.Text = "#PROJECT TESTERS EMAIL#"
.Replacement.Text = Tester
.Forward = True
.Wrap = wdFindContinue & vcCrLf
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
'''''''''''''''







End Sub
