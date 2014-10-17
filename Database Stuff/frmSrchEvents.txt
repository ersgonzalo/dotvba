Option Compare Database

Private Sub btnSrchClr_Click()
    txtNmeSrch = ""
    txtLocSrch = ""
    txtDateSrch = ""
    Me.txtNmeSrch.SetFocus
End Sub

Private Function DQ(s As Variant) As String
' double-up double quotes for SQL
DQ = Replace(Nz(s, ""), """", """""", 1, -1, vbBinaryCompare)
End Function

Private Sub btnSrch_Click()
Me.lstSrchRes.SetFocus
    'Clears current Listbox of any values
    Me.lstSrchRes.Value = Null
    'Used to populate listbox using an SQL statement.
    'This function also takes nulls/empty values for the form.
    Me.lstSrchRes.RowSource = _
        "SELECT eventsID , eventName, Description, CreatedBy, eventDate FROM tblEvents " & _
                "WHERE eventName LIKE ""*" & DQ(Me.txtNmeSrch.Value) & _
                    "*"" AND eventLocation LIKE ""*" & DQ(Me.txtLocSrch.Value) & _
                        "*"" AND eventDate LIKE ""*" & DQ(Me.txtDateSrch.Value) & "*"""
    
End Sub

Private Sub btnViewEvent_Click()
    
    'Checks to make sure that the Listbox value is not blank and then opens the viewable form.
    If Me.lstSrchRes.Value <> "" Then
        DoCmd.OpenForm "frmEventsView", , , "[eventsID]=" & Me.lstSrchRes.Column(0)
    Else
        MsgBox "Please make a search first or choose one of the listed events!"
    End If
    
End Sub

Private Sub lstSrchRes_DblClick(Cancel As Integer)
    
     'Checks to make sure that the Listbox value is not blank and then opens the viewable form.
    If Me.lstSrchRes.Value <> "" Then
        DoCmd.OpenForm "frmEventsView", , , "[eventsID]=" & Me.lstSrchRes.Column(0)
    Else
        MsgBox "Please make a search first!"
    End If
    
End Sub

Private Sub btnEditEvent_Click()
    
    'Checks to make sure that the Listbox value is not blank and then opens the editable form.
    If Me.lstSrchRes.Value <> "" Then
        DoCmd.OpenForm "frmEventsEdit", , , "[eventsID]=" & Me.lstSrchRes.Column(0)
    Else
        MsgBox "Please make a search or choose one of the listed events!"
    End If
        
End Sub

Private Sub Form_Load()

    If usrPerm = 3 Then
        btnEditEvent.Enabled = False
    End If
    Call UserCheck
    
End Sub
