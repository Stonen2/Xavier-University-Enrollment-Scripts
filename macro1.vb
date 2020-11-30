Sub InsertSignature()
'
' Print Decision Letter Macro
' This Macro is intended for all Regular admit letters sent out through Xavier. 
' This will automatically add Aaron Meis signature to each individual letter stored locally on undergraduate instance of slate.
' 
' Created by Nick Stone 9/25/17
' Updated by Jonathan Thortnon 9/26/18

Selection.Find.Execute Replace:=wdReplaceAll

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Sincerely,"
.Replacement.Text = "Sincerely,"

Do While .Execute
With Dialogs(wdDialogInsertPicture)
.Name = ""
Selection.TypeText Text:="Sincerely," & vbCrLf & vbCrLf




.Execute
End With
Loop
.Forward = True
.Wrap = wdFindContinue & vcCrLf
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
End Sub
