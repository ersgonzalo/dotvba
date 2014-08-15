'-----------------------------------------------------------------------'
'Code by Eric Gonzalo                                                   '
'For use in the Department of Transportation to help with mass printing '
'files from a single folder. However, it seems to go out of order no    '
'matter what. Helps with violations and other general work.             '
'-----------------------------------------------------------------------'

Option Explicit

Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) _
As Long

Public Sub PrintSetup(ByVal strPathAndFilename As String)
    'Method to print a file using the library from system
    Call apiShellExecute(Application.hwnd, "print", strPathAndFilename, vbNullString, vbNullString, 0)
     
End Sub

Sub PrintFiles()
    Dim FSO As New FileSystemObject
    Dim myFolder As Folder
    Dim myFile As file
    Dim fileN As String
    
    If Range("D5").Value = "" Then
        MsgBox "Please enter a folder you would like to print from!"
        Exit Sub
    Else
        Set myFolder = FSO.GetFolder(Range("D5").Value)
    End If

    If ((Sheets("Sheet1").OLEObjects("chkList").Object.Value = True)) Then
        'Debug.Print "Hello!"
        Range("A2").Select
        Do Until IsEmpty(ActiveCell)
        fileN = "?" & ActiveCell.Value & "*" & ".*"
            For Each myFile In myFolder.Files
                If myFile.Name Like fileN Then
                    'Method to print the found file in the origin folder
					PrintSetup (myFile)
					'Method will wait 7 seconds before printing the next file.
                    Application.Wait (Now + TimeValue("00:00:07"))
                End If
            Next
            ActiveCell.Offset(1, 0).Select
        Loop
    Else
        For Each myFile In myFolder.Files
                 'To see if such stuff exist:: Debug.Print myFile.Name & " in " & myFile.Path 'Or do whatever you want with the file                 
                 PrintSetup (myFile)
                 'Debug.Print myFile & "...next File"
                 Application.Wait (Now + TimeValue("00:00:07"))
        Next
    
    End If
            
    MsgBox "Your files should have all been printed! Please check to make sure there is no mistakes or unusual copies!"
    
End Sub
