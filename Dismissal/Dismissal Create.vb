'-----------------------------------------------------------------------'
'Code by Eric Gonzalo                                                   '
'Used in the Department of Transportation to help creating Dismissal    '
'Inspections from property data. Still in progress but should work with '
'the data and create a specific sheet based on the row that gets filled '
'out.                                                                   '
'-----------------------------------------------------------------------'
Option Explicit

Sub DismissMake()

On Error GoTo errorHandler

Dim wdApp As Word.Application
Dim myDoc As Word.Document
Dim mywdRange As Word.Range
Dim tempLoc As String
Dim savLoc As String
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

    If Sheets("Introduction").Range("B4").Value = "" Then
        MsgBox "Please enter a location for your template file!"
        Exit Sub
    Else
        tempLoc = Sheets("Introduction").Range("B4")
    End If
    
    If Sheets("Introduction").Range("B6").Value = "" Then
        MsgBox "Please enter a name for your Desktop Folder!"
        Exit Sub
    Else
        savLoc = "C:\Users\" & (Environ$("Username")) & "\Desktop\" & Sheets("Introduction").Range("B6").Value & "\"
        
        If Dir(savLoc, vbDirectory) = "" Then
        'Checks to see if folder exists otherwise will create such a folder.
        'Will end up preventing Dismissal from being created if MkDir does not have this check.
            MkDir savLoc
        End If
    End If
    
    Set myDoc = wdApp.Documents.Add(Template:=tempLoc)
    Set d_Borough = Sheets("Creation").Range("A2")
    Set d_Address = Sheets("Creation").Range("B2")
    Set d_Block = Sheets("Creation").Range("C2")
    Set d_Lot = Sheets("Creation").Range("D2")
    Set d_Vio = Sheets("Creation").Range("E2")
    Set d_Permit = Sheets("Creation").Range("F2")
    Set d_Attempt = Sheets("Creation").Range("G2")

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
    
    With wdApp.ActiveDocument
        .SaveAs savLoc & d_Borough & " - " & d_Address
        .Application.Quit
    End With
    
    'Should I just insert a method to go down excel rows instead of deleting?
    'This one works more easily, and it allows a user to check it instead of being stuck with a loop.
    Rows(2).Delete

'MsgBox "You file have been created! Please check to make sure all the data is present"

errorHandler:
Set wdApp = Nothing
Set myDoc = Nothing
Set mywdRange = Nothing

End Sub

Sub FolderOpen()
    
    Dim savLoc As String
    
    If Sheets("Introduction").Range("B6").Value = "" Then
        MsgBox "Please let the Transfer Function create your Desktop folder first!"
        Exit Sub
    Else
        savLoc = "C:\Users\" & (Environ$("Username")) & "\Desktop\" & Sheets("Introduction").Range("B6").Value & "\"
    End If

    Call Shell("explorer.exe" & " " & savLoc, vbNormalFocus)
    
End Sub
