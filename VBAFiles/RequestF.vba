Option Compare Database
Private Sub ACRegCb_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType.Value = rs("ACType").Value
            Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator.Value = rs("Operator").Value
            Me.Customer.Value = rs("Customer").Value
        Else
            Me.ACType.Value = Null
            Me.ACMTOW.Value = Null
            Me.Operator.Value = Null
            Me.Customer.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ACType.Value = Null
        Me.ACMTOW.Value = Null
        Me.Operator.Value = Null
        Me.Customer.Value = Null
    End If

 Exit Sub

End Sub
Private Sub AddFlightBtn_Click()
    ' Check If all the required fields are filled
    If IsNull(Me.RefNo.Value) Or IsNull(Me.ReqNo.Value) Or IsNull(Me.ReqDate.Value) Or IsNull(Me.Customer.Value) Or IsNull(Me.Operator.Value) _
        Or IsNull(Me.ACType.Value) Or IsNull(Me.ACMTOW.Value) Or IsNull(Me.ACRegCb.Value) Or IsNull(Me.ORIGIN.Value) Or IsNull(Me.DEST.Value) Or IsNull(Me.PFL.Value) Or IsNull(Me.Schedule.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    End If

    ' Add a New record To the table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("NewFlightRequestsT", dbOpenDynaset)

    rs.AddNew

    ' Assign values To the fields in the table
    With rs
        .Fields("RefNo").Value = Me.RefNo.Value
        .Fields("ReqNo").Value = Me.ReqNo.Value
        .Fields("ReqDate").Value = Me.ReqDate.Value
        .Fields("Customer").Value = Me.Customer.Value
        .Fields("Operator").Value = Me.Operator.Value
        .Fields("ACType").Value = Me.ACType.Value
        .Fields("ACMTOW").Value = Me.ACMTOW.Value
        .Fields("ACReg").Value = Me.ACReg.Value
        .Fields("PFL").Value = Me.PFL.Value
        .Fields("Schedule").Value = Me.Schedule.Value
        .Fields("ORIGIN").Value = Me.ORIGIN.Value
        .Fields("DEST").Value = Me.DEST.Value
        ' Save the New record
        .Update
    End With

    ' Clean up

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset form values


    ' Display a message indicating successful addition
    MsgBox "Data added successfully."
    Me.Undo
End Sub
Private Sub newRequestBtn_Click()
    ' Clear the form fields To prepare For a New record
    ' Set the default value For the RefNo field

    cmdGenerate_Click
    ' Set focus To the first field For data entry
    Me.ReqNo.SetFocus
End Sub
Private Sub cmdGenerate_Click()
    Dim RefValue As Integer
    Dim rst As DAO.Recordset

    RefValue = Nz(DLookup("RefNum", "TRefNo"), 0)

    Me.RefNo = "TAS-" & Format(RefValue + 1, "33120")

    Set rst = CurrentDb.OpenRecordset("TRefNo", dbOpenDynaset, dbSeeChanges)

    With rst
        .Edit
        ![RefNum] = ![RefNum] + 1
        .Update
    End With

    rst.Close
    Set rst = Nothing

End Sub

Private Function DataLoad()
    ' Check If the form is Not in DataEntry mode
    If Me.NewRecord Then
        ' Get the latest record from NewFlightRequestsT table
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim strSQL As String

        Set db = CurrentDb
        strSQL = "Select TOP 1 * FROM NewFlightRequestsT ORDER BY AddedTime DESC"
        Set rs = db.OpenRecordset(strSQL)

        ' Check If a record was found
        If Not rs.EOF Then
            ' Assign the field values To the form controls
            Me.RefNo.Value = rs.Fields("RefNo").Value
            Me.ReqNo.Value = rs.Fields("ReqNo").Value
            Me.ReqDate.Value = rs.Fields("ReqDate").Value
            Me.Customer.Value = rs.Fields("Customer").Value
            Me.Operator.Value = rs.Fields("Operator").Value
            Me.ACType.Value = rs.Fields("ACType").Value
            Me.ACMTOW.Value = rs.Fields("ACMTOW").Value
            Me.ACReg.Value = rs.Fields("ACReg").Value
            Me.PFL.Value = rs.Fields("PFL").Value
            Me.Schedule.Value = rs.Fields("Schedule").Value
            Me.ORIGIN.Value = rs.Fields("ORIGIN").Value
            Me.DEST.Value = rs.Fields("DEST").Value
        End If

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
End Function

Private Sub ResetBtn_Click()
    Me.RefNo.Value = ""
    Me.ReqNo.Value = ""
    Me.ReqDate.Value = ""
    Me.Customer.Value = ""
    Me.Operator.Value = ""
    Me.ACType.Value = ""
    Me.ACMTOW.Value = ""
    Me.ACRegCb.Value = ""
    Me.PFL.Value = ""
    Me.Schedule.Value = ""
    Me.ORIGIN.Value = ""
    Me.DEST.Value = ""
    Me.Undo
    ' DoCmd.cancelEvent
    ' DoCmd.RunCommand acCmdRecordsGoToNew
End Sub

Private Sub updateFlightbtn_Click()
    DataLoad
End Sub




DoCmd.RunCommand acCmdSaveRecord
DoCmd.RunCommand acCmdRecordsGoToNew