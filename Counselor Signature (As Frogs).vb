Sub InsertSignatureLastTry()
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
.Name = "https://images.mentalfloss.com/sites/default/files/styles/mf_image_16x9/public/502534-iStock-153768983.jpg?itok=a4zItlvW&resize=1100x1100"







.Execute
Selection.TypeText Text:="Rachel Richter" & vbCrLf
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
.Name = "http://naturemappingfoundation.org/natmap/photos/amphibians/american_bullfrog_0184np.jpg"





.Execute
Selection.TypeText Text:="Brian Gipson" & vbCrLf & vbCrLf
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
.Name = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQkT6ykLtM24ZxMm6zFmBd_0ugnlYeRTBn3l8TPrdA0FDkOZDvGDQ"





.Execute
Selection.TypeText Text:="Arianna Dunn" & vbCrLf & vbCrLf
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

.Name = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT-5BFKT7ca0yeByQ7iQ0NM5KZsmpSl0orC-sYvc7AD2ohjOLOZpQ"
Selection.TypeText Text:=" " & vbCrLf & vbCrLf




.Execute
Selection.TypeText Text:="Ashley Zeller" & vbCrLf & vbCrLf
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
.Name = "https://i.kinja-img.com/gawker-media/image/upload/s--pGLnA03a--/c_scale,f_auto,fl_progressive,q_80,w_800/bqhumfnolbpcz4drj4vy.jpg"





.Execute
Selection.TypeText Text:="Chloe Storm" & vbCrLf & vbCrLf
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
.Name = "https://pmdvod.nationalgeographic.com/NG_Video/306/151/smpost_1510772094465.jpg"





.Execute
Selection.TypeText Text:="David Donnelly" & vbCrLf & vbCrLf
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
.Name = "https://static.independent.co.uk/s3fs-public/thumbnails/image/2018/06/14/12/frog-pond.jpg?w968h681"





.Execute
Selection.TypeText Text:="Erin Melody" & vbCrLf & vbCrLf
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
.Name = "jb20vaW1hZ2VzL2kvMDAwLzA5Ni8yODEvb3JpZ2luYWwvd2hpdGUtdHJlZS1mcm9nLmpwZw==https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRHohSag9Y5M5cLULNwpN9CIuK3v1cXX63_cVi1lMUvMM4G8wMv"





.Execute
Selection.TypeText Text:="Hannah Shirkey" & vbCrLf & vbCrLf
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
.Name = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTbNlkw_129iVOBFK3N8-duInY1cCBeLYxvsmIn4usa1b99FGo"





.Execute
Selection.TypeText Text:="Jay Cloutier" & vbCrLf & vbCrLf
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
.Name = "https://usercontent2.hubstatic.com/13746147_f496.jpg"





.Execute
Selection.TypeText Text:="Julie Nelson" & vbCrLf & vbCrLf
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
.Name = "https://images.unsplash.com/photo-1506506447188-78e2a1051d9e?ixlib=rb-0.3.5&ixid=eyJhcHBfaWQiOjEyMDd9&s=6f7806c0a01af0521ed3e8062ce137d2&w=1000&q=80"





.Execute
Selection.TypeText Text:="Kaitlin McGeeney" & vbCrLf & vbCrLf
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
.Name = "http://www.aquariumofpacific.org/images/made/images/uploads/20170502_AOPfrog_pacifictree_5123_900_600_80auto.jpg"





.Execute
Selection.TypeText Text:="Kelly Conklin" & vbCrLf & vbCrLf
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
.Name = "http://www.ontarioparks.com/parksblog/wp-content/uploads/2015/10/Bullfrog.jpg"




.Execute
Selection.TypeText Text:="Lindsey Steller" & vbCrLf & vbCrLf
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
.Name = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSJEZRrCjjJmerfuQaK9kOWg75-1FKD706LwABNu6SYxytfSakTQw"





.Execute
Selection.TypeText Text:="Mike Garcia" & vbCrLf & vbCrLf
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
.Name = "https://i.pinimg.com/originals/bf/4e/40/bf4e4067252227bd3f758bba7dcee2ff.jpg"





.Execute
Selection.TypeText Text:="Tim Wilmes" & vbCrLf & vbCrLf
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