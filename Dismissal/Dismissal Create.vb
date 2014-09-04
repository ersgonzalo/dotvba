'-----------------------------------------------------------------------'
'Code by Eric Gonzalo                                                   '
'Used in the Department of Transportation to help creating Dismissal    '
'Inspections from property data. Still in progress but should work with '
'the data and create a specific sheet based on the row that gets filled '
'out. Still needs a way to check the data and save it.                  '
'-----------------------------------------------------------------------'
Option Explicit

Sub CertGenerator()

On Error GoTo errorHandler

Dim wdApp As Word.Application
Dim myDoc As Word.Document
Dim mywdRange As Word.Range
Dim d_Borough As Excel.Range 'For Bookmarklets
    Dim boro As String
Dim d_Address As Excel.Range
Dim d_Block As Excel.Range
Dim d_Lot As Excel.Range
Dim d_Vio As Excel.Range
Dim d_Permit As Excel.Range
Dim d_Attempt As Excel.Range 'End of Bookmarklets

Set wdApp = New Word.Application

    With wdApp
        .Visible = True
        .WindowState = wdWindowStateMaximize
    End With

    Set myDoc = wdApp.Documents.Add(Template:="C:\Users\EGonzalo\Downloads\Dismissal Template.dotx")
    Set d_Borough = Sheets("Information").Range("A2")
    Set d_Address = Sheets("Information").Range("B2")
    Set d_Block = Sheets("Information").Range("C2")
    Set d_Lot = Sheets("Information").Range("D2")
    Set d_Vio = Sheets("Information").Range("E2")
    Set d_Permit = Sheets("Information").Range("F2")
    Set d_Attempt = Sheets("Information").Range("G2")

    Select Case d_Borough
        Case "K"
            boro = "Brooklyn"
        Case "M"
            boro = "Manhattan"
        Case "Q"
            boro = "Queens"
        Case "R"
            boro = "Staten Island"
        Case "X"
            boro = "Bronx"
    End Select
    
    With myDoc.Bookmarks
        .Item("d_borough").Range.InsertAfter boro
        .Item("block").Range.InsertAfter d_Block
        .Item("lot").Range.InsertAfter d_Lot
        .Item("d_violation").Range.InsertAfter d_Vio
        .Item("d_permitNum").Range.InsertAfter d_Permit
        .Item("attemptNum").Range.InsertAfter d_Attempt
        .Item("d_vioAddr").Range.InsertAfter d_Address
    End With
    
'Insert method to go down excel rows?

    With wdApp.ActiveDocument
        .SaveAs d_Borough & " - " & d_Address
    End With

MsgBox "You file has been created!"

errorHandler:
Set wdApp = Nothing
Set myDoc = Nothing
Set mywdRange = Nothing

End Sub

' d_borough, block, lot, d_violation, d_permitNum, attemptNum, d_vioAddr
'http://www.mrexcel.com/forum/general-excel-discussion-other-questions/787952-best-way-populate-word.html
'http://www.mrexcel.com/forum/excel-questions/706989-export-excel-ranges-word-bookmarks-using-visual-basic-applications-save-word-doc-same-location-workbook.html
'http://www.mrexcel.com/forum/excel-questions/544782-visual-basic-applications-code-excel-opens-word-template-fills-bookmarks-but-cant-get-percent-formatting-right.html
