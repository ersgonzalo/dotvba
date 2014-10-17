Option Compare Database
Option Explicit

Private Sub btnLoginExit_Click()
    
    'Exits the database from the login form.
    DoCmd.Quit
    
End Sub

Private Sub cboEmployee_AfterUpdate()
    'After selecting user name set focus to password field
    Me.txtPassword.SetFocus
End Sub

Private Sub cmdLogin_Click()

Dim lngMyEmpID As String
Dim intLogonAttempts As Double
    
    'Check to see if data is entered into the UserName combo box
    If IsNull(Me.cboEmployee) Or Me.cboEmployee = "" Then
      MsgBox "You must select a Login ID.", vbOKOnly, "Required Data"
        Me.cboEmployee.SetFocus
        Exit Sub
    End If

    'Check to see if data is entered into the password box
    If IsNull(Me.txtPassword) Or Me.txtPassword = "" Then
      MsgBox "You must enter a Password.", vbOKOnly, "Required Data"
        Me.txtPassword.SetFocus
        Exit Sub
    End If

    'Check value of password in tblEmployees to see if this
    'matches value chosen in combo box
    If Me.txtPassword.Value = DLookup("EmployeePassword", "tblLogin", _
            "[loginEmpID]=" & Me.cboEmployee.Value) Then

        lngMyEmpID = Me.cboEmployee.Value
        usrNameID = lngMyEmpID
        Call UserDefine

        'Close login form and open calendar main form
        DoCmd.Close acForm, "frmLogin", acSaveNo
        DoCmd.OpenForm "frmCalendar"

        Else
            MsgBox "Your password is invalid. Please try again!", vbOKOnly, _
                "Invalid Entry!"
            Me.txtPassword.SetFocus
    End If

    'If User Enters incorrect password 5 times database will shutdown
    intLogonAttempts = intLogonAttempts + 1
    If intLogonAttempts > 5 Then
      MsgBox "You have tried too many times. Please contact an Admin for help.", _
               vbCritical, "Restricted Access!"
        Application.Quit
    End If

End Sub


