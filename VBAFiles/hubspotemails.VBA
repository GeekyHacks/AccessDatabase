Option Compare Database

Private Sub ACRegCb1_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb1.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType1.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator1.Value = rs("Operator").Value
            Me.Customer1.Value = rs("Customer").Value
        Else
            Me.ACType1.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator1.Value = Null
            Me.Customer1.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ACType1.Value = Null
        'Me.ACMTOW.Value = Null
        Me.Operator1.Value = Null
        Me.Customer1.Value = Null
    End If

 Exit Sub

End Sub


Private Sub ACRegCb2_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb2.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType2.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator2.Value = rs("Operator").Value
            Me.Customer2.Value = rs("Customer").Value
        Else
            Me.ACType2.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator2.Value = Null
            Me.Customer2.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb3_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb3.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType3.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator3.Value = rs("Operator").Value
            Me.Customer3.Value = rs("Customer").Value
        Else
            Me.ACType3.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator3.Value = Null
            Me.Customer3.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb4_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb4.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType4.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator4.Value = rs("Operator").Value
            Me.Customer4.Value = rs("Customer").Value
        Else
            Me.ACType4.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator4.Value = Null
            Me.Customer4.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb5_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb2.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType5.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator5.Value = rs("Operator").Value
            Me.Customer5.Value = rs("Customer").Value
        Else
            Me.ACType5.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator5.Value = Null
            Me.Customer5.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb6_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb2.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType6.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator6.Value = rs("Operator").Value
            Me.Customer6.Value = rs("Customer").Value
        Else
            Me.ACType6.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator6.Value = Null
            Me.Customer6.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub
Private Sub UpdateUser1()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User1 = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
Private Sub UpdateUser2()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User2 = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
Private Sub UpdateUser3()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User3 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub
Private Sub UpdateUser4()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User4 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub
Private Sub UpdateUser5()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User5 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub
Private Sub UpdateUser6()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User6 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub

Private Sub ActualOPSDate1_Click()
    InputDateField ActualOPSDate1, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate2_Click()
    InputDateField ActualOPSDate2, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate3_Click()
    InputDateField ActualOPSDate3, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate4_Click()
    InputDateField ActualOPSDate4, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate5_Click()
    InputDateField ActualOPSDate5, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate6_Click()
    InputDateField ActualOPSDate6, "Select a date To use this on your form"
End Sub

Private Sub addOPSDateBtn_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim I As Integer
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID

    LoggedUser = DLookup("UserName", "UsersT", "ID =" & UserID & "")
    Dim txtCallSign As String, OPSDate As String, ActualOPSDate As String, ORIGIN As String, DEST As String, RefNo As Variant, PermitNo As String, ACRegCb As String, ACTyp As String, Operator As String, Customer As String, User As String
    ' Open the main table's recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("OVRF_OpsT", dbOpenDynaset)

    ' Loop through the sets of text boxes And add records
    For I = 1 To 6 ' Change 5 To the desired number of OPSDATE

        ' Retrieve values from text boxes
        txtCallSign = Nz(Me.Controls("txtCallSign" & I).Value, "")
        OPSDate = Nz(Me.Controls("OPSDate" & I).Value, "")
        ActualOPSDate = Nz(Me.Controls("ActualOPSDate" & I).Value, "")
        ORIGIN = Nz(Me.Controls("ORIGIN" & I).Value, "")
        DEST = Nz(Me.Controls("DEST" & I).Value, "")
        RefNo = Nz(Me.Controls("RefNo" & I).Value, "")
        'PermitNo = Nz(Me.Controls("PermitNo" & I).Value, "")
        ACRegCb = Nz(Me.Controls("ACRegCb" & I).Value, "")
        ACType = Nz(Me.Controls("ACType" & I).Value, "")
        Operator = Nz(Me.Controls("Operator" & I).Value, "")
        Customer = Nz(Me.Controls("Customer" & I).Value, "")
        User = Nz(Me.Controls("User" & I).Value, "")

        ' Check If at least one field is Not empty
        If txtCallSign <> "" Or OPSDate <> "" Or ActualOPSDate <> "" Or ORIGIN <> "" Or DEST <> "" Or RefNo <> "" Or PermitNo <> "" Or ACRegCb <> "" Or ACType <> "" Or Operator <> "" Or Customer <> "" Or User <> "" Then
            ' Add a record
            rs.AddNew
            ' Assign values To the fields in the table
            rs.Fields("txtCallSign").Value = txtCallSign
            rs.Fields("OPSDate").Value = OPSDate
            rs.Fields("ActualOPSDate").Value = ActualOPSDate
            rs.Fields("ORIGIN").Value = ORIGIN
            rs.Fields("DEST").Value = DEST
            rs.Fields("RefNo").Value = RefNo
            'rs.Fields("OPSPermitNo").Value = PermitNo
            rs.Fields("ACReg").Value = ACRegCb
            rs.Fields("ACType").Value = ACType
            rs.Fields("Operator").Value = Operator
            rs.Fields("Customer").Value = Customer
            rs.Fields("User").Value = User
            rs.Update
        End If

        ' Clear the text boxes after adding a record
        Me.Controls("txtCallSign" & I).Value = Null
        Me.Controls("OPSDate" & I).Value = Null
         Me.Controls("ActualOPSDate" & I).Value = Null
        Me.Controls("ORIGIN" & I).Value = Null
        Me.Controls("DEST" & I).Value = Null
        Me.Controls("RefNo" & I).Value = Null
        'Me.Controls("PermitNo" & I).Value = Null
        'Me.Controls("CAMA").Value = Null
        Me.Controls("ACRegCb" & I).Value = Null
        Me.Controls("ACType" & I).Value = Null
        Me.Controls("Operator" & I).Value = Null
        Me.Controls("Customer" & I).Value = Null
        Me.Controls("User" & I).Value = Null
    Next I

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Optional: Display a message indicating successful addition
    MsgBox "Flights record added successfully.", vbInformation

 Exit Sub

End Sub

Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "OFR_NewAddOpsF", acSaveNo
End Sub
Private Sub Form_GotFocus()
 Form_Load
End Sub

Private Sub LoadLastBtn_Click()
    Me.CAMA = GetLastPermitNo()
    Me.LastRefNo = GetLastRefNo()
End Sub
Private Function GetLastPermitNo() As Variant

    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    'Dim lastRefNo As Variant
    Dim MaxID As Long
    MaxID = DMax("ID", "OVRF_RequestT")
    ' Me.RecordSource = "Select * FROM NewSectorsT WHERE ID = " & MaxID & ";"

    Set db = CurrentDb
    strSQL = "Select * FROM OVRF_RequestT WHERE ID = " & MaxID & ";"
    Set rs = db.OpenRecordset(strSQL)

    ' Check If a record was found
    If Not rs.EOF Then
        PermitNo = rs.Fields("PermitNo").Value
        LastRefNo = rs.Fields("RefNo").Value
    End If

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLastPermitNo = PermitNo
End Function
Private Function GetLastRefNo() As Variant

    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    'Dim lastRefNo As Variant
    Dim MaxID As Long
    MaxID = DMax("ID", "OVRF_RequestT")
    ' Me.RecordSource = "Select * FROM NewSectorsT WHERE ID = " & MaxID & ";"

    Set db = CurrentDb
    strSQL = "Select * FROM OVRF_RequestT WHERE ID = " & MaxID & ";"
    Set rs = db.OpenRecordset(strSQL)

    ' Check If a record was found
    If Not rs.EOF Then
        LastRefNo = rs.Fields("RefNo").Value
    End If

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLastRefNo = LastRefNo
End Function
Public Sub OPSDate2_GotFocus()
    UpdateUser2
    Me.OPSDate2 = Me.OPSDate1
End Sub
Public Sub OPSDate3_GotFocus()
    UpdateUser3
End Sub
Public Sub OPSDate4_GotFocus()
    UpdateUser4
    Me.OPSDate4 = Me.OPSDate3
End Sub

Public Sub OPSDate5_GotFocus()
    UpdateUser5
End Sub
Public Sub OPSDate6_GotFocus()
    UpdateUser6
    Me.OPSDate6 = Me.OPSDate5
End Sub
Public Sub OPSDate1_Click()
    UpdateUser1
    InputDateField OPSDate1, "Select a date To use this on your form"
  
End Sub
Public Sub OPSDate2_Click()
    InputDateField OPSDate2, "Select a date To use this on your form"
End Sub
Public Sub OPSDate3_Click()
    InputDateField OPSDate3, "Select a date To use this on your form"
End Sub
Public Sub OPSDate4_Click()
    InputDateField OPSDate4, "Select a date To use this on your form"
End Sub
Public Sub OPSDate5_Click()
    InputDateField OPSDate5, "Select a date To use this on your form"
End Sub
Public Sub OPSDate6_Click()
    InputDateField OPSDate6, "Select a date To use this on your form"
End Sub

Private Sub ORIGIN1_GotFocus()
  Me.ActualOPSDate1 = Me.OPSDate1
End Sub
Private Sub ORIGIN2_GotFocus()
    Me.ActualOPSDate2 = Me.OPSDate2
    Me.ORIGIN2 = Me.DEST1
End Sub

Private Sub DEST2_GotFocus()
    Me.DEST2 = Me.ORIGIN1
    Me.RefNo2 = Me.LastRefNo
    Me.ACRegCb2 = Me.ACRegCb1
    ACRegCb2_Change
End Sub
Private Sub ORIGIN3_GotFocus()
 Me.ActualOPSDate3 = Me.OPSDate3
    Me.ORIGIN3 = Me.DEST2
End Sub
Private Sub DEST3_GotFocus()
    Me.DEST3 = Me.ORIGIN2
    Me.RefNo3 = Me.LastRefNo
    Me.ACRegCb3 = Me.ACRegCb2
    ACRegCb3_Change
End Sub
Private Sub ORIGIN4_GotFocus()
 Me.ActualOPSDate4 = Me.OPSDate4
    Me.ORIGIN4 = Me.DEST3
End Sub
Private Sub DEST4_GotFocus()
    Me.DEST4 = Me.ORIGIN3
    Me.RefNo4 = Me.LastRefNo
    Me.ACRegCb4 = Me.ACRegCb3
    ACRegCb4_Change
End Sub
Private Sub ORIGIN5_GotFocus()
 Me.ActualOPSDate5 = Me.OPSDate5
    Me.ORIGIN5 = Me.DEST4
End Sub
Private Sub DEST5_GotFocus()
    Me.DEST5 = Me.ORIGIN4
    Me.RefNo6 = Me.LastRefNo
    Me.ACRegCb6 = Me.ACRegCb5
    ACRegCb6_Change
End Sub
Private Sub ORIGIN6_GotFocus()
 Me.ActualOPSDate6 = Me.OPSDate6
    Me.ORIGIN6 = Me.DEST5
End Sub
Private Sub DEST6_GotFocus()
    Me.DEST6 = Me.ORIGIN5
    Me.RefNo5 = Me.LastRefNo
    Me.ACRegCb5 = Me.ACRegCb4
    ACRegCb5_Change
End Sub
Public Sub ResetBtn_Click()
    On Error GoTo ErrorHandler
        Me.Controls("CAMA").Value = Null
        Me.Controls("LastRefNo").Value = Null
        For I = 1 To 6
            Me.Controls("txtCallSign" & I).Value = Null
            Me.Controls("OPSDate" & I).Value = Null
            Me.Controls("ActualOPSDate" & I).Value = Null
            Me.Controls("ORIGIN" & I).Value = Null
            Me.Controls("DEST" & I).Value = Null
            Me.Controls("RefNo" & I).Value = Null
            'Me.Controls("PermitNo" & I).Value = Null
            ' Me.Controls("CAMA").Value = Null
            Me.Controls("ACRegCb" & I).Value = Null
            Me.Controls("ACType" & I).Value = Null
            Me.Controls("Operator" & I).Value = Null
            Me.Controls("Customer" & I).Value = Null
            Me.Controls("User" & I).Value = Null
        Next I
        Me.Undo
        DoCmd.CancelEvent

     Exit Sub

ErrorHandler:
        MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
Dim CAMA As Variant
Dim LastRefNo As Variant
Dim OVR_FORM As Form

On Error Resume Next
Set OVR_FORM = Forms("Update_OVFP_FORM")
On Error GoTo 0

If OVR_FORM Is Nothing Then
    ' Form is not loaded
    LoadLastBtn_Click
    Me.closeBtn.Visible = False
    Me.ActualOPSDate1.Visible = False
    Me.ActualOPSDate2.Visible = False
    Me.ActualOPSDate3.Visible = False
    Me.ActualOPSDate4.Visible = False
    Me.ActualOPSDate5.Visible = False
    Me.ActualOPSDate6.Visible = False
Else
    Me.closeBtn.Visible = True
    Me.ActualOPSDate1.Visible = True
    Me.ActualOPSDate2.Visible = True
    Me.ActualOPSDate3.Visible = True
    Me.ActualOPSDate4.Visible = True
    Me.ActualOPSDate5.Visible = True
    Me.ActualOPSDate6.Visible = True
    ' Form is loaded
    If OVR_FORM.Visible = False Then
        LoadLastBtn_Click
    Else
        If OVR_FORM.Controls("Update_OVRF_RequestF").Form Is Nothing Then
            LoadLastBtn_Click
        Else
            CAMA = OVR_FORM.Controls("Update_OVRF_RequestF").Form.Controls("RefNo").Column(0)
            LastRefNo = OVR_FORM.Controls("Update_OVRF_RequestF").Form.Controls("txtRefNo")
            If IsNull(CAMA) Then
                LoadLastBtn_Click
            Else
                Me.CAMA = CAMA
                Me.LastRefNo = LastRefNo
            End If
        End If
    End If
End If
End Sub

Private Sub txtCallSign2_GotFocus()
    Me.txtCallSign2 = Me.txtCallSign1
End Sub

Private Sub txtCallSign3_GotFocus()
    Me.txtCallSign3 = Me.txtCallSign1
End Sub
Private Sub txtCallSign4_GotFocus()
    Me.txtCallSign4 = Me.txtCallSign2
End Sub
Private Sub txtCallSign5_GotFocus()
    Me.txtCallSign5 = Me.txtCallSign3
End Sub
Private Sub txtCallSign6_GotFocus()
    Me.txtCallSign6 = Me.txtCallSign4
End Sub




    ///////////////////////////////////////////////

    Option Compare Database

Private Sub ACRegCb1_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb1.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType1.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator1.Value = rs("Operator").Value
            Me.Customer1.Value = rs("Customer").Value
        Else
            Me.ACType1.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator1.Value = Null
            Me.Customer1.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ACType1.Value = Null
        'Me.ACMTOW.Value = Null
        Me.Operator1.Value = Null
        Me.Customer1.Value = Null
    End If

 Exit Sub

End Sub


Private Sub ACRegCb2_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb2.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType2.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator2.Value = rs("Operator").Value
            Me.Customer2.Value = rs("Customer").Value
        Else
            Me.ACType2.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator2.Value = Null
            Me.Customer2.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb3_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb3.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType3.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator3.Value = rs("Operator").Value
            Me.Customer3.Value = rs("Customer").Value
        Else
            Me.ACType3.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator3.Value = Null
            Me.Customer3.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb4_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb4.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType4.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator4.Value = rs("Operator").Value
            Me.Customer4.Value = rs("Customer").Value
        Else
            Me.ACType4.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator4.Value = Null
            Me.Customer4.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb5_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb2.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType5.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator5.Value = rs("Operator").Value
            Me.Customer5.Value = rs("Customer").Value
        Else
            Me.ACType5.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator5.Value = Null
            Me.Customer5.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub

Private Sub ACRegCb6_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb2.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType6.Value = rs("ACType").Value
            'Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator6.Value = rs("Operator").Value
            Me.Customer6.Value = rs("Customer").Value
        Else
            Me.ACType6.Value = Null
            'Me.ACMTOW.Value = Null
            Me.Operator6.Value = Null
            Me.Customer6.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        'Me.ACType2.Value = Null
        'Me.ACMTOW.Value = Null
        'Me.Operator2.Value = Null
        'Me.Customer2.Value = Null
     Exit Sub
    End If


End Sub
Private Sub UpdateUser1()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User1 = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
Private Sub UpdateUser2()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User2 = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
Private Sub UpdateUser3()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User3 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub
Private Sub UpdateUser4()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User4 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub
Private Sub UpdateUser5()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User5 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub
Private Sub UpdateUser6()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User6 = DLookup("UserName", "UsersT", "ID =" & UserID & "")

End Sub

Private Sub ActualOPSDate1_Click()
    InputDateField ActualOPSDate1, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate2_Click()
    InputDateField ActualOPSDate2, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate3_Click()
    InputDateField ActualOPSDate3, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate4_Click()
    InputDateField ActualOPSDate4, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate5_Click()
    InputDateField ActualOPSDate5, "Select a date To use this on your form"
End Sub
Private Sub ActualOPSDate6_Click()
    InputDateField ActualOPSDate6, "Select a date To use this on your form"
End Sub

Private Sub addOPSDateBtn_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim I As Integer
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID

    LoggedUser = DLookup("UserName", "UsersT", "ID =" & UserID & "")
    Dim txtCallSign As String, OPSDate As String, ActualOPSDate As String, ORIGIN As String, DEST As String, RefNo As Variant, PermitNo As String, ACRegCb As String, ACTyp As String, Operator As String, Customer As String, User As String
    ' Open the main table's recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("OVRF_OpsT", dbOpenDynaset)

    ' Loop through the sets of text boxes And add records
    For I = 1 To 6 ' Change 5 To the desired number of OPSDATE

        ' Retrieve values from text boxes
        txtCallSign = Nz(Me.Controls("txtCallSign" & I).Value, "")
        OPSDate = Nz(Me.Controls("OPSDate" & I).Value, "")
        ActualOPSDate = Nz(Me.Controls("ActualOPSDate" & I).Value, "")
        ORIGIN = Nz(Me.Controls("ORIGIN" & I).Value, "")
        DEST = Nz(Me.Controls("DEST" & I).Value, "")
        RefNo = Nz(Me.Controls("RefNo" & I).Value, "")
        PermitNo = Nz(Me.Controls("PermitNo" & I).Value, "")
        ACRegCb = Nz(Me.Controls("ACRegCb" & I).Value, "")
        ACType = Nz(Me.Controls("ACType" & I).Value, "")
        Operator = Nz(Me.Controls("Operator" & I).Value, "")
        Customer = Nz(Me.Controls("Customer" & I).Value, "")
        User = Nz(Me.Controls("User" & I).Value, "")

        ' Check If at least one field is Not empty
        If txtCallSign <> "" Or OPSDate <> "" Or ActualOPSDate <> "" Or ORIGIN <> "" Or DEST <> "" Or RefNo <> "" Or PermitNo <> "" Or ACRegCb <> "" Or ACType <> "" Or Operator <> "" Or Customer <> "" Or User <> "" Then
            ' Add a record
            rs.AddNew
            ' Assign values To the fields in the table
            rs.Fields("txtCallSign").Value = txtCallSign
            rs.Fields("OPSDate").Value = OPSDate
            rs.Fields("ActualOPSDate").Value = ActualOPSDate
            rs.Fields("ORIGIN").Value = ORIGIN
            rs.Fields("DEST").Value = DEST
            rs.Fields("RefNo").Value = RefNo
            'rs.Fields("OPSPermitNo").Value = PermitNo
            rs.Fields("ACReg").Value = ACRegCb
            rs.Fields("ACType").Value = ACType
            rs.Fields("Operator").Value = Operator
            rs.Fields("Customer").Value = Customer
            rs.Fields("User").Value = User
            rs.Update
        End If

        ' Clear the text boxes after adding a record
        Me.Controls("txtCallSign" & I).Value = Null
        Me.Controls("OPSDate" & I).Value = Null
         Me.Controls("ActualOPSDate" & I).Value = Null
        Me.Controls("ORIGIN" & I).Value = Null
        Me.Controls("DEST" & I).Value = Null
        Me.Controls("RefNo" & I).Value = Null
        'Me.Controls("PermitNo" & I).Value = Null
        ' Me.Controls("CAMA").Value = Null
        Me.Controls("ACRegCb" & I).Value = Null
        Me.Controls("ACType" & I).Value = Null
        Me.Controls("Operator" & I).Value = Null
        Me.Controls("Customer" & I).Value = Null
        Me.Controls("User" & I).Value = Null
    Next I

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Optional: Display a message indicating successful addition
    MsgBox "Flights record added successfully.", vbInformation

 Exit Sub

End Sub

Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "OFR_NewAddOpsF", acSaveNo
End Sub
Private Sub Form_GotFocus()
 Form_Load
End Sub

Private Sub LoadLastBtn_Click()
    Me.CAMA = GetLastPermitNo()
    Me.LastRefNo = GetLastRefNo()
End Sub
Private Function GetLastPermitNo() As Variant

    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    'Dim lastRefNo As Variant
    Dim MaxID As Long
    MaxID = DMax("ID", "OVRF_RequestT")
    ' Me.RecordSource = "Select * FROM NewSectorsT WHERE ID = " & MaxID & ";"

    Set db = CurrentDb
    strSQL = "Select * FROM OVRF_RequestT WHERE ID = " & MaxID & ";"
    Set rs = db.OpenRecordset(strSQL)

    ' Check If a record was found
    If Not rs.EOF Then
        PermitNo = rs.Fields("PermitNo").Value
        LastRefNo = rs.Fields("RefNo").Value
    End If

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLastPermitNo = PermitNo
End Function
Private Function GetLastRefNo() As Variant

    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    'Dim lastRefNo As Variant
    Dim MaxID As Long
    MaxID = DMax("ID", "OVRF_RequestT")
    ' Me.RecordSource = "Select * FROM NewSectorsT WHERE ID = " & MaxID & ";"

    Set db = CurrentDb
    strSQL = "Select * FROM OVRF_RequestT WHERE ID = " & MaxID & ";"
    Set rs = db.OpenRecordset(strSQL)

    ' Check If a record was found
    If Not rs.EOF Then
        LastRefNo = rs.Fields("RefNo").Value
    End If

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLastRefNo = LastRefNo
End Function
Public Sub OPSDate2_GotFocus()
    UpdateUser2
    Me.OPSDate2 = Me.OPSDate1
End Sub
Public Sub OPSDate3_GotFocus()
    UpdateUser3
End Sub
Public Sub OPSDate4_GotFocus()
    UpdateUser4
    Me.OPSDate4 = Me.OPSDate3
End Sub

Public Sub OPSDate5_GotFocus()
    UpdateUser5
End Sub
Public Sub OPSDate6_GotFocus()
    UpdateUser6
    Me.OPSDate6 = Me.OPSDate5
End Sub
Public Sub OPSDate1_Click()
    UpdateUser1
    InputDateField OPSDate1, "Select a date To use this on your form"
  
End Sub
Public Sub OPSDate2_Click()
    InputDateField OPSDate2, "Select a date To use this on your form"
End Sub
Public Sub OPSDate3_Click()
    InputDateField OPSDate3, "Select a date To use this on your form"
End Sub
Public Sub OPSDate4_Click()
    InputDateField OPSDate4, "Select a date To use this on your form"
End Sub
Public Sub OPSDate5_Click()
    InputDateField OPSDate5, "Select a date To use this on your form"
End Sub
Public Sub OPSDate6_Click()
    InputDateField OPSDate6, "Select a date To use this on your form"
End Sub
Private Sub PermitNo1_GotFocus()
    Me.PermitNo1 = Me.CAMA
    Me.RefNo1 = Me.LastRefNo
End Sub
Private Sub PermitNo2_GotFocus()
    Me.PermitNo2 = Me.PermitNo1
    Me.RefNo2 = Me.LastRefNo
    Me.ACRegCb2 = Me.ACRegCb1
    ACRegCb2_Change
End Sub
Private Sub PermitNo3_GotFocus()
    Me.PermitNo3 = Me.PermitNo1
    Me.RefNo3 = Me.LastRefNo
    Me.ACRegCb3 = Me.ACRegCb2
    ACRegCb3_Change
End Sub
Private Sub PermitNo4_GotFocus()
    Me.PermitNo4 = Me.PermitNo1
    Me.RefNo4 = Me.LastRefNo
    Me.ACRegCb4 = Me.ACRegCb3
    ACRegCb4_Change
End Sub
Private Sub PermitNo5_GotFocus()
    Me.PermitNo5 = Me.PermitNo1
    Me.RefNo5 = Me.LastRefNo
    Me.ACRegCb5 = Me.ACRegCb4
    ACRegCb5_Change
End Sub
Private Sub PermitNo6_GotFocus()
    Me.PermitNo6 = Me.PermitNo5
    Me.RefNo6 = Me.LastRefNo
    Me.ACRegCb6 = Me.ACRegCb5
    ACRegCb6_Change
End Sub
Private Sub ORIGIN1_GotFocus()
  Me.ActualOPSDate1 = Me.OPSDate1
End Sub
Private Sub ORIGIN2_GotFocus()
    Me.ActualOPSDate2 = Me.OPSDate2
    Me.ORIGIN2 = Me.DEST1
End Sub
Private Sub DEST2_GotFocus()
    Me.DEST2 = Me.ORIGIN1
End Sub
Private Sub ORIGIN3_GotFocus()
 Me.ActualOPSDate3 = Me.OPSDate3
    Me.ORIGIN3 = Me.DEST2
End Sub
Private Sub DEST3_GotFocus()
    Me.DEST3 = Me.ORIGIN2
End Sub
Private Sub ORIGIN4_GotFocus()
 Me.ActualOPSDate4 = Me.OPSDate4
    Me.ORIGIN4 = Me.DEST3
End Sub
Private Sub DEST4_GotFocus()
    Me.DEST4 = Me.ORIGIN3
End Sub
Private Sub ORIGIN5_GotFocus()
 Me.ActualOPSDate5 = Me.OPSDate5
    Me.ORIGIN5 = Me.DEST4
End Sub
Private Sub DEST5_GotFocus()
    Me.DEST5 = Me.ORIGIN4
End Sub
Private Sub ORIGIN6_GotFocus()
 Me.ActualOPSDate6 = Me.OPSDate6
    Me.ORIGIN6 = Me.DEST5
End Sub
Private Sub DEST6_GotFocus()
    Me.DEST6 = Me.ORIGIN5
End Sub
Public Sub ResetBtn_Click()
    On Error GoTo ErrorHandler
        Me.Controls("CAMA").Value = Null
        Me.Controls("LastRefNo").Value = Null
        For I = 1 To 6
            Me.Controls("txtCallSign" & I).Value = Null
            Me.Controls("OPSDate" & I).Value = Null
            Me.Controls("ActualOPSDate" & I).Value = Null
            Me.Controls("ORIGIN" & I).Value = Null
            Me.Controls("DEST" & I).Value = Null
            Me.Controls("RefNo" & I).Value = Null
            Me.Controls("PermitNo" & I).Value = Null
            ' Me.Controls("CAMA").Value = Null
            Me.Controls("ACRegCb" & I).Value = Null
            Me.Controls("ACType" & I).Value = Null
            Me.Controls("Operator" & I).Value = Null
            Me.Controls("Customer" & I).Value = Null
            Me.Controls("User" & I).Value = Null
        Next I
        Me.Undo
        DoCmd.CancelEvent

     Exit Sub

ErrorHandler:
        MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
Dim CAMA As Variant
Dim LastRefNo As Variant
Dim OVR_FORM As Form

On Error Resume Next
Set OVR_FORM = Forms("Update_OVFP_FORM")
On Error GoTo 0

If OVR_FORM Is Nothing Then
    ' Form is not loaded
    LoadLastBtn_Click
    Me.closeBtn.Visible = False
    Me.ActualOPSDate1.Visible = False
    Me.ActualOPSDate2.Visible = False
    Me.ActualOPSDate3.Visible = False
    Me.ActualOPSDate4.Visible = False
    Me.ActualOPSDate5.Visible = False
    Me.ActualOPSDate6.Visible = False
Else
    Me.closeBtn.Visible = True
    Me.ActualOPSDate1.Visible = True
    Me.ActualOPSDate2.Visible = True
    Me.ActualOPSDate3.Visible = True
    Me.ActualOPSDate4.Visible = True
    Me.ActualOPSDate5.Visible = True
    Me.ActualOPSDate6.Visible = True
    ' Form is loaded
    If OVR_FORM.Visible = False Then
        LoadLastBtn_Click
    Else
        If OVR_FORM.Controls("Update_OVRF_RequestF").Form Is Nothing Then
            LoadLastBtn_Click
        Else
            CAMA = OVR_FORM.Controls("Update_OVRF_RequestF").Form.Controls("RefNo").Column(0)
            LastRefNo = OVR_FORM.Controls("Update_OVRF_RequestF").Form.Controls("txtRefNo")
            If IsNull(CAMA) Then
                LoadLastBtn_Click
            Else
                Me.CAMA = CAMA
                Me.LastRefNo = LastRefNo
            End If
        End If
    End If
End If
End Sub

Private Sub txtCallSign2_GotFocus()
    Me.txtCallSign2 = Me.txtCallSign1
End Sub

Private Sub txtCallSign3_GotFocus()
    Me.txtCallSign3 = Me.txtCallSign1
End Sub
Private Sub txtCallSign4_GotFocus()
    Me.txtCallSign4 = Me.txtCallSign2
End Sub
Private Sub txtCallSign5_GotFocus()
    Me.txtCallSign5 = Me.txtCallSign3
End Sub
Private Sub txtCallSign6_GotFocus()
    Me.txtCallSign6 = Me.txtCallSign4
End Sub

