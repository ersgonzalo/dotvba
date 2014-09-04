'-----------------------------------------------------------------------'
'Code by Eric Gonzalo                                                   '
'Used in the Department of Transportation to help with mass printing 	'
'files from a single folder. However, it seems to go out of order when  '
'printer slows in processing, likely due to the network. Helps with 	'
'violations and other general documents. Built on a principle similar	'
'to the File Copier I had heavily researched previously.				'
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
    Dim copyCount As Integer
    Dim endCount As Integer
    
    If Range("D5").Value = "" Then
        'Checks to define a folder you print files from.
        MsgBox "Please enter a folder you would like to print from!"
        Exit Sub
    Else
        Set myFolder = FSO.GetFolder(Range("D5").Value)
    End If

    If Range("D11").Value = "" Or 1 Then
		'Helps define the number of copies that will be printed.
        Range("D11").Value = 1
        Range("E11").Value = "Copy"
    Else
        Range("E11").Value = "Copies"
    End If
    
    endCount = Range("D11").Value
    
    If ((Sheets("Sheet1").OLEObjects("chkList").Object.Value = True)) Then
        Range("A2").Select
        
        Do Until IsEmpty(ActiveCell)
            fileN = "?" & ActiveCell.Value & "*" & ".*"
        
            For Each myFile In myFolder.Files
            
                If myFile.Name Like fileN Then
                    For copyCount = 1 To endCount
                        PrintSetup (myFile)
                        'Debug.Print myFile & "...next File"
                    Next copyCount
                    'Method will wait 7 seconds before printing the next file.
                    Application.Wait (Now + TimeValue("00:00:07"))
                End If
                
            Next
            
        ActiveCell.Offset(1, 0).Select
        Loop
    Else
        For Each myFile In myFolder.Files
            For copyCount = 1 To endCount
                PrintSetup (myFile)
                'To see if such stuff exist:: Debug.Print myFile.Name & " in " & myFile.Path 'Or do whatever you want with the file
            Next copyCount
        Application.Wait (Now + TimeValue("00:00:07"))
        Next
    End If
            
    MsgBox "Your files should have all been printed! Please check to make sure there is no mistakes or unusual copies!"
    
End Sub