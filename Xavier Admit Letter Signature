Sub InsertSignature()
Selection.Find.Execute Replace:=wdReplaceAll

Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Sincerely,"
.Replacement.Text = "Sincerely,"

Do While .Execute
With Dialogs(wdDialogInsertPicture)
.Name = "*****"
Selection.TypeText Text:="Sincerely," & vbCrLf & " " & vbCrLf & " "




.Execute
End With
Selection.TypeText Text:="" & vbCrLf & " "
Loop
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
End Sub
