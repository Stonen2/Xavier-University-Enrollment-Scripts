Public Sub ExecutiveSummary()
Dim objExcel As Object


Dim find As String
Dim Title As String

Dim des As String
Dim reme As String


Dim cve As String
Dim sev As String


Dim bus As String
Dim Tech As String





Set objExcel = CreateObject("Excel.Application")
Dim exWb As Object
Dim b As String
Dim i As Integer
Dim found As Boolean
found = False
Dim max As Integer



Set exWb = objExcel.Workbooks.Open("C:\Users\Admin\Desktop\Testing VB Scriots\w.xlsx")
i = 2
While found = False

find = exWb.Worksheets(1).Range("A" & i)

If find = "" Then

found = True


End If

i = i + 1



Wend

max = i - 2

' MsgBox (max)

    Dim myRange As Range
    find = exWb.Worksheets(1).Range("A1")
    Set myRange = ActiveDocument.Range.Bookmarks("Bookmark1").Range
    ActiveDocument.Tables.Add Range:=myRange, NumRows:=max, NumColumns:=8
   ' ActiveDocument.Tables(1).Cell(1, 1).Range = "Test"
    'ActiveDocument.Tables(1).Cell(2, 2).Range = Find
    Dim count As Integer
    count = 0
    Dim track As Integer
   
    Dim bb As Integer
    bb = 2
    While count < max
    find = exWb.Worksheets(1).Range("A" & bb)
    Title = exWb.Worksheets(1).Range("B" & bb)
    des = exWb.Worksheets(1).Range("C" & bb)
    reme = exWb.Worksheets(1).Range("D" & bb)
    cve = exWb.Worksheets(1).Range("E" & bb)
    sev = exWb.Worksheets(1).Range("F" & bb)
    bus = exWb.Worksheets(1).Range("G" & bb)
    Tech = exWb.Worksheets(1).Range("H" & bb)
   
    ActiveDocument.Tables(1).Cell(track, 1).Range = find
    ActiveDocument.Tables(1).Cell(track, 2).Range = Title
    ActiveDocument.Tables(1).Cell(track, 3).Range = des
    ActiveDocument.Tables(1).Cell(track, 4).Range = reme
    ActiveDocument.Tables(1).Cell(track, 5).Range = cve
    ActiveDocument.Tables(1).Cell(track, 6).Range = sev
    ActiveDocument.Tables(1).Cell(track, 7).Range = bus
    ActiveDocument.Tables(1).Cell(track, 8).Range = Tech
   '  MsgBox (Find)
   
    count = count + 1
    track = track + 1
   
    bb = bb + 1
    Wend
   
   

' Make the Tables in Excel



'Fill the table contents
exWb.Close





End Sub
