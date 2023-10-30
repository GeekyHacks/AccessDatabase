Option Compare Database

Private Sub ACRegCb_Change()
    On Error GoTo ErrorHandler

        Dim selectedValue As Variant
        selectedValue = Me.ACRegCb.Value ' Get the selected value from the combo box

        If Not IsNull(selectedValue) Then
            Me.ACType.Value = DLookup("ACType", "ACRegT", "ACReg = '" & selectedValue & "'")
            Me.ACMTOW.Value = DLookup("ACMTOW", "ACRegT", "ACReg = '" & selectedValue & "'")
        Else
            Me.ACType.Value = Null
            Me.ACMTOW.Value = Null
        End If

     Exit Sub

ErrorHandler:
        MsgBox "Error: " & Err.Description, vbExclamation, "Error"
End Sub
Private Sub closeBtn_Click()
    ' Check If all the required fields are empty, including RefNo
    If IsNull(Me.ReqDate.Value) And IsNull(Me.Customer.Value) And IsNull(Me.Operator.Value) And IsNull(Me.ACType.Value) And IsNull(Me.ACMTOW.Value) And IsNull(Me.ACRegCb.Value) And IsNull(Me.PFL.Value) And IsNull(Me.Schedule.Value) And IsNull(Me.RefNo.Value) Then
        ' If all the required fields are empty, including RefNo, close the form without adding a record
        DoCmd.Close acForm, "FlightRequestF"
    Else
        ' If the form is Not empty, check If there are unsaved changes
        If Me.Dirty Then
            ' If the form has unsaved changes, prompt the user To save the changes
            If MsgBox("Do you want To save the changes?", vbQuestion + vbYesNo, "Save Changes") = vbYes Then
            If MsgBox("Are you sure, this will create a duplicate fligth request?", vbQuestion + vbYesNo, "Save Changes") = vbYes Then
            DoCmd.Save
            End If
            
            End If
        End If

        ' Close the form
        DoCmd.Close acForm, "FlightRequestF"
    End If
End Sub
Private Sub AddFlightBtn_Click()
    ' Check if all the required fields are filled
    If IsNull(Me.ReqDate.Value) Or IsNull(Me.Customer.Value) Or IsNull(Me.Operator.Value) _
        Or IsNull(Me.ACType.Value) Or IsNull(Me.ACMTOW.Value) Or IsNull(Me.ACRegCb.Value) _
        Or IsNull(Me.PFL.Value) Or IsNull(Me.Schedule.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
        Exit Sub
    End If

    ' Add a new record to the table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("FlightRequestsT", dbOpenDynaset)

    rs.AddNew

    ' Assign values to the fields in the table
    rs.Fields("RefNo").Value = Me.RefNo.Value
    rs.Fields("ReqNo").Value = Me.ReqNo.Value
    rs.Fields("ReqDate").Value = Me.ReqDate.Value
    rs.Fields("Customer").Value = Me.Customer.Value
    rs.Fields("Operator").Value = Me.Operator.Value
    rs.Fields("ACType").Value = Me.ACType.Value
    rs.Fields("ACMTOW").Value = Me.ACMTOW.Value
    rs.Fields("ACReg").Value = Me.ACReg.Value
    rs.Fields("PFL").Value = Me.PFL.Value
    rs.Fields("Schedule").Value = Me.Schedule.Value

    ' Save the new record
    On Error GoTo ErrorHandler
    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset form values
    Me.ReqNo.Value = ""
    Me.ReqDate.Value = Null
    Me.Customer.Value = ""
    Me.Operator.Value = ""
    Me.ACType.Value = ""
    Me.ACMTOW.Value = Null
    Me.ACRegCb.Value = ""
    Me.PFL.Value = ""
    Me.Schedule.Value = ""

    ' Display a message indicating successful addition
    MsgBox "Data added successfully."
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
    rs.CancelUpdate
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

Private Sub addSectorBtn_Click()
    DoCmd.OpenForm "AddSectorF", WindowMode:=acDialog
End Sub

Private Sub newRequestBtn_Click()
    ' Check if there are any unsaved changes
    If Me.Dirty Then
        Me.Undo ' Undo any changes to cancel the current record
    End If
    
    ' Clear the form fields to prepare for a new record
    Me.RefNo.Value = GenerateNewRefNo()
    Me.ReqNo.Value = ""
    Me.ReqDate.Value = Null
    Me.Customer.Value = ""
    Me.Operator.Value = ""
    Me.ACType.Value = ""
    Me.ACMTOW.Value = Null
    Me.ACReg.Value = ""
    Me.PFL.Value = ""
    Me.Schedule.Value = ""

    ' Set focus to the first field for data entry
    Me.ReqNo.SetFocus
End Sub

Private Function GenerateNewRefNo() As String
    ' Calculate the New RefNo value based on existing records
    Dim lastRefNo As Variant
    lastRefNo = DMax("RefNo", "FlightRequestsT")

    Dim newRefNo As String

    If IsNull(lastRefNo) Then
        ' If the table is empty, set the initial reference number
        newRefNo = "TAS-000001"
    Else
        ' Generate a random number between 1 and 999999
        Dim randomNumber As Long
        Randomize  ' Initialize the random number generator
        randomNumber = Int((999999 - 1 + 1) * Rnd + 1)

        ' Combine random number with the prefix and format it
        newRefNo = "TAS" & Format(randomNumber, "000000")
    End If

    ' Return the new RefNo value
    GenerateNewRefNo = newRefNo
End Function

Private Sub updateFlightbtn_Click()
    DoCmd.Save
End Sub
Private Sub loadForm()
    ' Check if the form is in DataEntry mode
    If Me.NewRecord Then
        ' Get the latest AddedTime value from FlightRequestsT table
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim strSQL As String
        Dim latestAddedTime As Variant

        Set db = CurrentDb
        strSQL = "SELECT MAX(AddedTime) AS LatestAddedTime FROM FlightRequestsT"
        Set rs = db.OpenRecordset(strSQL)

        ' Check if there is a latest AddedTime value
        If Not rs.EOF Then
            latestAddedTime = rs.Fields("AddedTime").Value

            ' Retrieve the record with the latest AddedTime
            strSQL = "SELECT TOP 1 * FROM FlightRequestsT WHERE AddedTime = #" & Format(latestAddedTime, "mm/dd/yyyy hh:mm:ss") & "# ORDER BY AddedTime DESC"
            rs.Close
            Set rs = db.OpenRecordset(strSQL)

            ' Check if a record was found
            If Not rs.EOF Then
                ' Assign the field values to the form controls
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
            End If
        End If

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
End Sub
Private Sub Form_Load()
   loadForm
End Sub
