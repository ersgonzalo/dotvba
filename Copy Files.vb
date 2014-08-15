'-----------------------------------------------------------------------'
'Code by Eric Gonzalo                                                   '
'For use in the Department of Transportation to help with moving files  '
'from place to place. Helps with violations and possibly some other     '
'types but it is slow if there are many folders in a location.          '
'-----------------------------------------------------------------------'

Sub FileManage()
      
      Dim folderName As String
      Dim boro As String
      Dim searchLoc As String
      Dim destnLoc As String 'specific locations
      Dim fileN As String
      
      boro = LCase(Range("D5").Value) 'Reads the cells, this one converts it to lower case.
      folderName = (Range("D7").Value)
      searchLoc = (Range("D9").Value)
      
      'Part 1 - Checks to make sure boro is not blank.
      If (boro = "") Then
        MsgBox "You have not typed-in/selected a borough!"
        Exit Sub
      End If
      
      'Part 2 - Uses the inputted folder folder name from the cell and then creates a folder with that name.
      If (folderName <> "") Then
            
            'Part 3 - Uses the specified path of the Shared Drive/whatever folder you are using.
            If (searchLoc <> "") Then
                searchLoc = searchLoc & "\" & Application.Proper(boro)
            Else
                MsgBox "You have not defined a location to search from!"
                Exit Sub
            End If
        'part 2 continued
        If Dir("C:\Users\" & (Environ$("Username")) & "\Desktop\" & folderName, vbDirectory) = "" Then
            MkDir ("C:\Users\" & (Environ$("Username")) & "\Desktop\" & folderName)
        Else
            MsgBox "The folder already exists, please check to make sure your folder name is not wrong or rename your folder!"
            Exit Sub
        End If
      Else
        MsgBox "Your folder requires a name!"
        Exit Sub
      End If
      
      'Part 4 - Copies the files.
      ' Select cell A2, *first line of data*.
      Range("A2").Select
      ' Set Do loop to stop when an empty cell is reached.
      Do Until IsEmpty(ActiveCell)
         'Search Method
         Call FileSearch(searchLoc)
         ' Step down 1 row from present location.
         ActiveCell.Offset(1, 0).Select
      Loop
      
      Call Shell("explorer.exe" & " " & "C:\Users\" & (Environ$("Username")) & "\Desktop\" & folderName, vbNormalFocus)
      MsgBox "The violations have been copied! Please double check the resulting folder to make sure no violations were missed."
      
End Sub

Function FileSearch(sPath As String) As String

    Dim FSO As New FileSystemObject
    Dim myFolder As Folder
    Dim mySubFolder As Folder
    Dim myFile As file
    Dim destnLoc As String
    
    fileN = "?" & ActiveCell.Value & "*" & ".pdf"
    destinLoc = "C:\Users\" & (Environ$("Username")) & "\Desktop\" & Range("D7").Value & "\"
         
    Set myFolder = FSO.GetFolder(sPath)
    
    For Each myFile In myFolder.Files
            If myFile.Name Like fileN Then
                 'To see if such stuff exist Debug.Print myFile.Name & " in " & myFile.Path 'Or do whatever you want with the file
                 'Method to copy the found file
                 FSO.CopyFile myFile, destinLoc
            End If
    Next
    
    
    For Each mySubFolder In myFolder.SubFolders
        For Each myFile In mySubFolder.Files
            If myFile.Name Like fileN Then
                 'To see if such stuff exist Debug.Print myFile.Name & " in " & myFile.Path 'Or do whatever you want with the file
                 'Method to copy the found file
                 FSO.CopyFile myFile, destinLoc
            End If
        Next
        FileSearch = FileSearch(mySubFolder.Path)
    Next
    
End Function

'Reconsiderations
'-----------------
'
'End Notes
'------------
'For searching subdirectories
'http://www.ammara.com/access_image_faq/recursive_folder_search.html
'https://stackoverflow.com/questions/20687810/vba-macro-that-search-for-file-in-multiple-subfolders