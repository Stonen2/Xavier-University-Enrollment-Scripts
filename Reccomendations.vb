Public Sub Reccomendations()
 
 'Set variables in order to access excel applications
 Dim objExcel As Object
Set objExcel = CreateObject("Excel.Application")
Dim exWb As Object
Dim b As String
Dim i As Integer
Dim found As Boolean
found = False
Dim max As Integer
Dim find As String

'This is the path of the document... This can be changed on a by user basis. Simple stuff really
Set exWb = objExcel.Workbooks.Open("C:\Users\Admin\Desktop\Testing VB Scriots\w.xlsx")
i = 2
'Find the last row in the table
While found = False

find = exWb.Worksheets(1).Range("A" & i)

If find = "" Then

found = True


End If

i = i + 1



Wend

max = i - 2

'MsgBox (max)
 'Make a table that is the length of the excel spreadseheet
    Dim myRange As Range
    find = exWb.Worksheets(1).Range("A1")
    Set myRange = ActiveDocument.Range.Bookmarks("Bookmark2").Range
    ActiveDocument.Tables.Add Range:=myRange, NumRows:=max, NumColumns:=4
 
 
 
 
 '''''''''''''''''''''''''''''''''''
 ' Assign some variables to pull the data
 
 Dim Title As String
 Dim execfind As String
 Dim execrec As String
 
 'Need an offset since excel starts at 1
  Dim count As Integer
    count = 0
    Dim track As Integer
   
    Dim bb As Integer
    bb = 2
   
    While count < max
    'Assign the dynamic variables
    find = exWb.Worksheets(1).Range("A" & bb)
    Title = exWb.Worksheets(1).Range("B" & bb)
    execfind = exWb.Worksheets(1).Range("L" & bb)
    execrec = exWb.Worksheets(1).Range("M" & bb)
    'Populate the Cells with the data above
    ActiveDocument.Tables(2).Cell(track, 1).Range = find
    ActiveDocument.Tables(2).Cell(track, 2).Range = Title
    ActiveDocument.Tables(2).Cell(track, 3).Range = execfind
    ActiveDocument.Tables(2).Cell(track, 4).Range = execrec

   '  MsgBox (Find)
   
    count = count + 1
    track = track + 1
   
    bb = bb + 1
    Wend
 
 
 ''''''''''''''''''''''''''''''''''''''
 exWb.Close
 
 'Finish
 
 End Sub
 
 Public Sub run()
 ' Written by Nick Stone 11/12/2019
 ' This program is designed to be run on the VulnTemplateMASTER document make sure you copy the document each time you run the program
 ' This is very much a functional program... OOP is not an option in visual basic
 ' This is the main of the module! Every program needs a main.
 ' No additional downloads needed just need any microsoft word document
 
 'Run the function to parse all of the first page data
 Call TestMacro
 'Add the first table in the exec summary
 Call ExecutiveSummary
 'Add the second table to the rec. in Exec summary
 Call Reccomendations
 
 
 'Finish
 
 End Sub
