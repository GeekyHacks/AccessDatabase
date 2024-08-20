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
    ' Check if all the required fields are filled
    If IsNull(Me.RefNo.Value) Or IsNull(Me.ReqNo.Value) Or IsNull(Me.ReqDate.Value) Or IsNull(Me.Customer.Value) Or IsNull(Me.Operator.Value) _
        Or IsNull(Me.ACType.Value) Or IsNull(Me.ACMTOW.Value) Or IsNull(Me.ACRegCb.Value) Or IsNull(Me.PFL.Value) Or IsNull(Me.Schedule.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
        Exit Sub
    End If

    ' Add a new record to the table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("NewFlightRequestsT", dbOpenDynaset)

    rs.AddNew

    ' Assign values to the fields in the table
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
        
        ' Save the new record
        .Update
    End With
    
    ' Clean up

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset form values
    Me.RefNo.Value = ""
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

End Sub

' Private Sub Form_AfterInsert()

' ' Check If all the required fields are empty, including RefNo
' If IsNull(Me.ReqDate.Value) And IsNull(Me.Customer.Value) And IsNull(Me.Operator.Value) And IsNull(Me.ACType.Value) And IsNull(Me.ACMTOW.Value) And IsNull(Me.ACRegCb.Value) And IsNull(Me.PFL.Value) And IsNull(Me.Schedule.Value) And IsNull(Me.RefNo.Value) Then
'     ' If all the required fields are empty, including RefNo, close the form without adding a record
'     DoCmd.Close acForm, "LASTFORM"
' Else
'     ' If the form is Not empty, check If there are unsaved changes
'     If Me.Dirty Then
'         ' If the form has unsaved changes, prompt the user To save the changes
'         If MsgBox("Do you want To save the changes?", vbQuestion + vbYesNo, "Save Changes") = vbYes Then
'             If MsgBox("Are you sure, this will create a duplicate flight request?", vbQuestion + vbYesNo, "Save Changes") = vbYes Then
'             Me.Undo
'                 DoCmd.Save
'             End If
'         End If
'     End If

'     ' Close the form
'     DoCmd.Close acForm, "LASTFORM"
' End If
' End Sub



Private Sub newRequestBtn_Click()

    ' Clear the form fields to prepare for a new record
        ' Set the default value for the RefNo field
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
    lastRefNo = DMax("RefNo", "NewFlightRequestsT")

    Dim newRefNo As String

    If IsNull(lastRefNo) Then
        ' If the table is empty, set the initial reference number
        newRefNo = "TAS-000001"
    Else
        ' Extract the numeric part of the last reference number
        Dim lastNumber As Long
        lastNumber = CLng(Mid(lastRefNo, 5))

        ' Generate the next number greater than the last number
        Dim nextNumber As Long
        nextNumber = lastNumber + 1

        ' Combine the next number with the prefix and format it
        newRefNo = "TAS-" & Format(nextNumber, "000000")
    End If

    ' Return the new RefNo value
    GenerateNewRefNo = newRefNo
End Function

' Private Sub updateFlightbtn_Click()
'    Form_Load
' End Sub

' Private Sub Form_Load()
'    ' Check if the form is not in DataEntry mode
'     If Not Me.NewRecord Then
'         ' Get the latest record from NewFlightRequestsT table
'         Dim db As DAO.Database
'         Dim rs As DAO.Recordset
'         Dim strSQL As String

'         Set db = CurrentDb
'         strSQL = "SELECT TOP 1 * FROM NewFlightRequestsT ORDER BY AddedTime DESC"
'         Set rs = db.OpenRecordset(strSQL)

'         ' Check if a record was found
'         If Not rs.EOF Then
'             ' Assign the field values to the form controls
'             Me.RefNo.Value = rs.Fields("RefNo").Value
'             Me.ReqNo.Value = rs.Fields("ReqNo").Value
'             Me.ReqDate.Value = rs.Fields("ReqDate").Value
'             Me.Customer.Value = rs.Fields("Customer").Value
'             Me.Operator.Value = rs.Fields("Operator").Value
'             Me.ACType.Value = rs.Fields("ACType").Value
'             Me.ACMTOW.Value = rs.Fields("ACMTOW").Value
'             Me.ACReg.Value = rs.Fields("ACReg").Value
'             Me.PFL.Value = rs.Fields("PFL").Value
'             Me.Schedule.Value = rs.Fields("Schedule").Value
'         End If

'         ' Clean up
'         rs.Close
'         Set rs = Nothing
'         Set db = Nothing
'     End If
' End Sub







    Private Sub Form_Load()
   ' Check if the form is not in DataEntry mode
    If Not Me.NewRecord Then
        ' Get the latest record from NewFlightRequestsT table
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim strSQL As String

        Set db = CurrentDb
        strSQL = "SELECT TOP 1 * FROM NewFlightRequestsT ORDER BY AddedTime DESC"
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

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
End Sub


    Private Sub closebtn_Click()
        ' Undo changes in subform1
           Me.RequestF.Form.Recordset.Undo
           
           ' Undo changes in subform2
           Me.NewAddSectorF.Form.Recordset.Undo
           
           ' Undo changes in subform3
           Me.NewMultiServicesF.Form.Recordset.Undo
           
           ' Close the main form
           DoCmd.Close acForm, Me.Name
       
       
               If Me.Dirty Then
                   ' If the form has unsaved changes, prompt the user To save the changes
                   If MsgBox("Do you want To save the changes?", vbQuestion + vbYesNo, "Save Changes") = vbYes Then
                   If MsgBox("Are you sure, this will create a duplicate fligth request?", vbQuestion + vbYesNo, "Save Changes") = vbYes Then
                   DoCmd.Save
                   End If
                   
                   End If
               End If
       
               ' Close the form
               DoCmd.Close acForm, "LASTFORM"
       
       
       End Sub
       