Option Compare Database

'Stores the login information
Public usrNameID As Integer
Public usrLogin As String
Public usrName As String
Public usrPerm As String
Public usrPass As String 'May not be safe, but is the most straightforward way to change passwords in frmLoginEdit

Public Function UserDefine()

Dim permConv As Integer

'Used to check the userNameID and then gets the User's Name associated with it
usrLogin = DLookup("EmployeeName", "tblLogin", "[loginEmpID]=" & [usrNameID])
usrPass = DLookup("EmployeePassword", "tblLogin", "[loginEmpID]=" & [usrNameID])
usrName = DLookup("personName", "tblPersons", "[loginEmpID]=" & [usrNameID])
usrPerm = DLookup("Permissions", "tblLogin", "[loginEmpID]=" & [usrNameID])

End Function

Public Function UserCheck()

    'Checks the current session to make sure that person is logged in.
    If usrNameID = 0 Then
     MsgBox "Session Expired. Please login again!", vbOKOnly, _
                "Login again!"
    DoCmd.Close
    DoCmd.OpenForm "frmLogin", View
    End If

End Function

Public Function EventUpdater()

Dim rs As DAO.Recordset

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [Tbl-Users] WHERE PKEnginnerID =" & sUser)
    If Not rs.EOF Then
        rs.Edit
        rs.Fields("fldLoggedIn").Value = False
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
    
End Function
    

'Old Practice Code for UserDefine()
'Dim db As Database
'Dim rs As DAO.Recordset
'Dim strSQL As String

'Replace with the SQL you need. Maybenot
'strSQL = "SELECT * FROM tblPersons WHERE loginID = """ & usrNameID & """;"
'Set db = CurrentDb
'Set rs = db.OpenRecordset(strSQL)

'Replace myFieldName with the appropriate field name
'usrName = rs.Fields("personName")

'rs.Close
