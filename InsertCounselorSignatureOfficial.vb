Sub InsertCounselorSignaturesOfficial()


'Created By Nick Stone 11/26/2018
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Rachel Richter"
.Replacement.Text = "Rachel Richter"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/RachelRichter_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/RachelRitcher_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/RachelRichter_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Rachel Richter" & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Brian Gipson"
.Replacement.Text = "Brian Gipson"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/BrianGipson_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/BrianGipson_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/BrianGipson_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Brian Gipson" & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Arianna Dunn"
.Replacement.Text = "Arianna Dunn"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/AriannaDunn_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/AriannaDunn_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/AriannaDunn_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Arianna Dunn" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Ashley Zeller"
.Replacement.Text = "Ashley Zeller"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/AshleyZeller_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/AshleyZeller_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/AshleyZeller_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Ashley Zeller" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Chloe Storm"
.Replacement.Text = "Chloe Storm"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/ChloeStorm_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/ChloeStorm_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/ChloeStorm_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Chloe Storm" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "David Donnelly"
.Replacement.Text = "David Donnelly"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/DavidDonnelly_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/DavidDonnelly_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/DavidDonnelly_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "David Donnelly" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Erin Melody"
.Replacement.Text = "Erin Melody"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/ErinMelody_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/ErinMelody_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/ErinMelody_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Erin Melody" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Hannah Shirkey"
.Replacement.Text = "Hannah Shirkey"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/HannahShirkey_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/HannahShirkey_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/HannahShirkey_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Hannah Shirkey" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Jay Cloutier"
.Replacement.Text = "Jay Cloutier"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/JayCloutier_signature.png
.Name = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTbNlkw_129iVOBFK3N8-duInY1cCBeLYxvsmIn4usa1b99FGo"
'.Name = "https://admissions.xavier.edu/www/images/JayCloutier_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Jay Cloutier" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Julie Nelson"
.Replacement.Text = "Julie Nelson"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/JulieNelson_signature.png
.Name = "https://usercontent2.hubstatic.com/13746147_f496.jpg"
'.Name = "https://admissions.xavier.edu/www/images/JulieNelson_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Julie Nelson" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Kaitlin McGeeney"
.Replacement.Text = "Kaitlin McGeeney"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/KaitlinMcGeeney_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/KaitlinMcGeeney_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/KaitlinMcGeeney_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Kaitlin McGeeney" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Kelly Conklin"
.Replacement.Text = "Kelly Conklin"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/KellyConklin_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/KellyConklin_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/KellyConklin_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Kelly Conklin" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Lindsey Steller"
.Replacement.Text = "Lindsey Steller"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/LindseySteller_signature.png"
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/LindseyStellar_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/LindseySteller_signature.png"
.Execute
Selection.TypeText Text:=vbCrLf & "Lindsey Steller" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Mike Garcia"
.Replacement.Text = "Mike Garcia"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/MikeGarcia_signature.png"
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/MichaelGarcia_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/MikeGarcia_signature.png"
.Execute
Selection.TypeText Text:=vbCrLf & "Mike Garcia" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.Find.Execute Replace:=wdReplaceAll
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.Text = "Tim Wilmes"
.Replacement.Text = "Tim Wilmes"
Do While .Execute
With Dialogs(wdDialogInsertPicture)
Selection.TypeText Text:=" " & vbCrLf & vbCrLf
'.Name = "https://admissions.xavier.edu/www/images/TimWilmes_signature.png
.Name = "https://admissions.xavier.edu/www/images/Counselor%20Signatures/TimWilmes_signature.png"
'.Name = "https://admissions.xavier.edu/www/images/TimWilmes_signature.png
.Execute
Selection.TypeText Text:=vbCrLf & "Tim Wilmes" & vbCrLf & vbCrLf
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub







