Option Compare Database

Private Sub btnCreateUser_Click()
On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb()
    Set rs = db.OpenRecordset(Name:="tblLogin", Options:=dbSeeChanges)
    Dim strSQL As String


    'Defines the permissions of a user.
    If cboUserPerm = "Administrator" Then
        permNum = 1
    ElseIf cboUserPerm = "Regular" Then
        permNum = 2
    ElseIf cboUserPerm = "Viewer" Then
        permNum = 3
    Else
        permNum = 3
    End If
    
    'Check to see if data is entered into the LoginID box
    If IsNull(Me.txtEmployeeName) Or Me.txtEmployeeName = "" Then
      MsgBox "You must fill in a Login ID.", vbOKOnly, "Required Data"
        Me.txtEmployeeName.SetFocus
        Exit Sub
    End If

    'Check to see if data is entered into the Password box
    If IsNull(Me.txtEmployeePassword) Or Me.txtEmployeePassword = "" Then
      MsgBox "You must enter a Password.", vbOKOnly, "Required Data"
        Me.txtEmployeePassword.SetFocus
        Exit Sub
    End If
    
    'Check to see if data is entered into the Password box
    If IsNull(Me.txtPasswordConfirm) Or Me.txtPasswordConfirm = "" Then
      MsgBox "You must retype your Password.", vbOKOnly, "Required Data"
        Me.txtPasswordConfirm.SetFocus
        Exit Sub
    End If
        
    'Checks to make sure your passwords matches, then adds it into the database.
    If Me.txtEmployeePassword = Me.txtPasswordConfirm Then
    With CurrentDb.OpenRecordset("tblLogin")
        .AddNew
            !EmployeeName = Me.txtEmployeeName
            !EmployeePassword = Me.txtEmployeePassword
            !Permissions = permNum
        .Update
'Must find a way to take the new LoginID and place it as loginempID in SQL statement below:
    End With
    'DoCmd.RunSQL "UPDATE tblPersons SET loginempID ='39' WHERE (tblPersons.personName = Me.cboPerson.Value);"
    Me.Requery
    MsgBox "Login account for " & Me.txtEmployeeName & " successfully created!"
        txtEmployeeName = ""
        txtEmployeePassword = ""
        txtPasswordConfirm = ""
        Exit Sub
    'Clears your entry if the passwords do not match.
    Else
        MsgBox "Your password does not match.", vbOKOnly, "Wrong Password"
        txtEmployeePassword = ""
        txtPasswordConfirm = ""
        Exit Sub
    End If
    
ExitSub:
    Exit Sub
ErrorHandler:
    'This is a roundabout way to fix the duplicates within the Login table. If you can
    'think of another way to do it within the function then do it!
    MsgBox "Please use another Login name!", , "Error"
    Resume ExitSub
End Sub

Private Sub btnLgnErase_Click()
    
    'Clears the values within the form.
    txtEmployeeName = ""
    txtEmployeePassword = ""
    txtPasswordConfirm = ""
    txtEmail = ""
    Me.txtEmployeeName.SetFocus
    
End Sub

Private Sub Form_Load()

    Call UserCheck
    
End Sub
