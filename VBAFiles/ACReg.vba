Option Compare Database

Private Sub closeBtn_Click()
DoCmd.Close acForm, "ACRegF"
End Sub


Private Sub AddACBtn_Click()
    ' Assign values to the fields in the form
    Me.Customer.Value = Nz(Me.Customer.Value, "")
    Me.Operator.Value = Nz(Me.Operator.Value, "")
    Me.ACType.Value = Nz(Me.ACType.Value, "")
    Me.ACMTOW.Value = Nz(Me.ACMTOW.Value, "")
    Me.ACReg.Value = Nz(Me.ACReg.Value, "")
    
    ' Add a new record to the table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("ACRegT", dbOpenDynaset)
    
    rs.AddNew
    
    ' Assign values to the fields in the table
    rs.Fields("Customer").Value = Me.Customer.Value
    rs.Fields("Operator").Value = Me.Operator.Value
    rs.Fields("ACType").Value = Me.ACType.Value
    rs.Fields("ACMTOW").Value = Me.ACMTOW.Value
    rs.Fields("ACReg").Value = Me.ACReg.Value
    
    ' Save the new record
    rs.Update
    
    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' Reset form values (excluding RefNo)
    Me.Customer.Value = ""
    Me.Operator.Value = ""
    Me.ACType.Value = ""
    Me.ACMTOW.Value = Null
    Me.ACReg.Value = ""
    
    ' Display a message indicating successful addition
        MsgBox "Data added successfully."
End Sub
Private Sub Form_Load()
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub UpdateACBtn_Click()
DoCmd.Save
DoCmd.Requery
End Sub

