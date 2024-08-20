Option Compare Database

Private Sub AddInvoice_Click()
    AddFuelInvoice
End Sub
Private Sub CustomerName_AfterUpdate()
    Dim CurrentCustomer As String
    Dim strSQL As String

    CurrentCustomer = Me.CustomerName.Value

    ' Use a function to build the query safely
    strSQL = BuildSafeQuery(CurrentCustomer)

    ' Set the RowSource of the combo box to the dynamic query
    Me.RefNo.RowSource = strSQL
    Me.RefNo.Requery

    ' Build the SQL query for OperatorListCb combo box
    strSQL = BuildSafeQueryForOperator(CurrentCustomer)
    Me.OperatorListCb.RowSource = strSQL
    Me.OperatorListCb.Requery
End Sub

Private Function BuildSafeQuery(ByVal CustomerValue As String) As String
    ' Build the SQL query using parameterized approach
    BuildSafeQuery = "SELECT DISTINCT RefNo, ID FROM FlightSupport_FuelPreformaQ_New_Test WHERE Customer = '" & Replace(CustomerValue, "'", "''") & "'"
End Function

Private Function BuildSafeQueryForOperator(ByVal CustomerValue As String) As String
    ' Build the SQL query using parameterized approach
    BuildSafeQueryForOperator = "SELECT DISTINCT Operator FROM FlightSupport_FuelPreformaQ_New_Test WHERE Customer = '" & Replace(CustomerValue, "'", "''") & "'"
End Function
Private Sub DueDate_Click()
    InputDateField DueDate, "Select a date To use this on your form"
End Sub

Private Sub InvoiceReset()
   'Me.RefNo.Value = Null
    Me.RefNo_Airport.Value = Null
    Me.Destination.Value = Null
    Me.Taxes.Value = Null
    Me.ServiceCost.Value = Null
    Me.Vendor.Value = Null
    Me.TripRefNo.Value = Null
    Me.DeliveryDate.Value = Null
    Me.Validity.Value = Null
    Me.Currency.Value = Null
    Me.ACType.Value = Null
    Me.Operator.Value = Null
    Me.Customer.Value = Null
    Me.ACReg.Value = Null
    Me.Amount.Value = Null
    Me.VATRate.Value = Null
    Me.Invoiced.Value = Null
    Me.InvoiceRefNo.Value = Null
    Me.VendorInvoice.Value = Null
    Me.InvoiceDate.Value = Null
    Me.DueDate.Value = Null
    Me.FuelStatus.Value = Null
    resetRowSource
End Sub

Private Sub EndDate_Change()
EndDate_AfterUpdate
End Sub

Private Sub EndDate_DblClick(Cancel As Integer)
EndDate_AfterUpdate
End Sub

Private Sub EndDate_GotFocus()
InputDateField EndDate, "Select a date To use this on your form"
End Sub

Private Sub InvoiceDate_Click()
    InputDateField InvoiceDate, "Select a date To use this on your form"
End Sub
Private Sub OperatorListCb_AfterUpdate()
    Dim CurrentOperator As String
    Dim strSQL As String

    CurrentOperator = Me.OperatorListCb.Value

    ' Use a Function To build the query safely
    strSQL = BuildSafeRefNoQueryFormOperator(CurrentOperator)

    ' Set the RowSource of the combo box To the dynamic query
    Me.RefNo.RowSource = strSQL
    Me.RefNo.Requery
End Sub

Private Function BuildSafeRefNoQueryFormOperator(ByVal OperatorValue As String) As String
    ' Build the SQL query using parameterized approach
    BuildSafeRefNoQueryFormOperator = "Select DISTINCT RefNo, ID FROM FlightSupport_FuelPreformaQ_New_Test WHERE Operator = '" & Replace(OperatorValue, "'", "''") & "'"
End Function

Private Sub RefNo_AfterUpdate()
    Dim RecordQuery As String
    Dim CurrentRefNO As Integer ' Declare CurrentRefNO As Variant

    CurrentRefNO = Me.RefNo.Column(1)

    RecordQuery = "Select Airport, Location_ID FROM FlightSupport_FuelPreformaQ_New_Test WHERE ID = " & CurrentRefNO

    ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    Me.RefNo_Airport.RowSource = RecordQuery

    ' Requery the AIRPORT combo box To reflect the updated options
    Me.RefNo_Airport.Requery

End Sub

Private Sub RefNo_Airport_AfterUpdate()
    RenderRecord
End Sub

Private Sub ResetBtn_Click()
    Me.RefNo.Value = Null
    Me.RefNo_Airport.Value = Null
    Me.Destination.Value = Null
    Me.Taxes.Value = Null
    Me.ServiceCost.Value = Null
    Me.Vendor.Value = Null
    Me.TripRefNo.Value = Null
    Me.DeliveryDate.Value = Null
    Me.Validity.Value = Null
    Me.Currency.Value = Null
    Me.ACType.Value = Null
    Me.Operator.Value = Null
    Me.Customer.Value = Null
    Me.ACReg.Value = Null
    Me.Amount.Value = Null
    Me.VATRate.Value = Null
    Me.Invoiced.Value = Null
    Me.InvoiceRefNo.Value = Null
    Me.VendorInvoice.Value = Null
    Me.InvoiceDate.Value = Null
    Me.DueDate.Value = Null
    Me.FuelStatus.Value = Null
    Me.CustomerName.Value = Null
    Me.OperatorListCb.Value = Null
    Me.ServiceStatus.Value = Null
    Me.StartDate.Value = Null
    Me.EndDate.Value = Null
    resetRowSource
End Sub

Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FlightSupport_EstimatedDebitF", acSaveNo
End Sub
Private Sub RenderRecord()
    Dim selectedAirport As Integer
    Dim rs As DAO.Recordset
    Dim db As DAO.Database

    selectedAirport = Me.RefNo_Airport.Column(1)

    If Not IsNull(selectedAirport) Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset("FlightSupport_FuelPreformaQ_New_Test", dbOpenSnapshot)

        ' Find the row that matches the selectedAirport And FlightType
        rs.FindFirst "Location_ID = " & selectedAirport

        If Not rs.NoMatch Then
            ' Assign values To the fields in the form
            Me.Customer.Value = rs("Customer").Value
            Me.Operator.Value = rs("Operator").Value
            Me.ACReg.Value = rs("ACReg").Value
            Me.ACType.Value = rs("ACType").Value
            Me.DeliveryDate.Value = rs("DepartDate").Value
            Me.TripRefNo.Value = rs("CallSign").Value
            Me.Vendor.Value = rs("Vendor").Value
            Me.Validity.Value = rs("Validity").Value
            Me.Currency.Value = rs("Currency").Value
            Me.ServiceCost.Value = rs("TASPrice").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.Destination.Value = rs("NextDest").Value
            Me.FuelStatus.Value = rs("FuelStatus").Value
            rs.Close
            Set rs = Nothing
            Set db = Nothing
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    Else
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
    End If
End Sub
Private Sub AddFuelInvoice()
    'On Error GoTo ErrorHandler
    Dim selectedAirport As Integer
    Dim selectedLocation As Variant
    Dim selectedRefNo As Variant
    Dim rs As DAO.Recordset
    Dim db As DAO.Database

    selectedRefNo = Me.RefNo.Column(0)
    selectedAirport = Me.RefNo_Airport.Column(1)
    selectedLocation = Me.RefNo_Airport.Column(0)
    Set db = CurrentDb
    Set rs = db.OpenRecordset("FlightSupport_EstimatedDebitT", dbOpenDynaset) ' Open in dbOpenDynaset mode

    ' Find the row that matches the selectedAirport And FlightType
    rs.FindFirst "Location_ID = " & selectedAirport
    If Not rs.NoMatch Then
        MsgBox "Record already exists."
        Exit Sub
    End If

    If IsNull(Me.RefNo) Or IsNull(Me.ACReg) Or IsNull(Me.Amount) Or IsNull(Me.VATRate) Or IsNull(Me.TotalCost) Or IsNull(Me.VAT) Or IsNull(Me.TotalGross) Or IsNull(Me.Vendor) Or IsNull(Me.TripRefNo) Or IsNull(Me.DeliveryDate) Or IsNull(Me.ACType) Or IsNull(Me.Operator) Or IsNull(Me.Customer) Or IsNull(Me.RefNo_Airport) Or IsNull(Me.Destination) Or IsNull(Me.Currency) Or IsNull(Me.ServiceCost) Or IsNull(Me.Validity) Then
        MsgBox "Add missing fields " & err.Description, vbCritical, "Error"
        Exit Sub
    Else
        rs.AddNew
        ' Assign values To the fields in the table
        rs.Fields("CallSign").Value = Me.TripRefNo.Value
        rs.Fields("Location").Value = selectedLocation
        rs.Fields("Location_ID").Value = selectedAirport
        rs.Fields("Destination").Value = Me.Destination.Value
        rs.Fields("Taxes").Value = Me.Taxes.Value
        rs.Fields("ServiceCost").Value = Me.ServiceCost.Value
        rs.Fields("Vendor").Value = Me.Vendor.Value
        rs.Fields("TripRefNo").Value = selectedRefNo
        rs.Fields("DeliveryDate").Value = Me.DeliveryDate.Value
        rs.Fields("Validity").Value = Me.Validity.Value
        rs.Fields("Currency").Value = Me.Currency.Value
        rs.Fields("ACType").Value = Me.ACType.Value
        rs.Fields("Operator").Value = Me.Operator.Value
        rs.Fields("Customer").Value = Me.Customer.Value
        rs.Fields("ACReg").Value = Me.ACReg.Value
        rs.Fields("Amount").Value = Me.Amount.Value
        rs.Fields("VATRate").Value = Me.VATRate.Value
        rs.Fields("Invoiced").Value = Me.Invoiced.Value
        rs.Fields("VendorRefNo").Value = Me.VendorInvoice.Value
        rs.Fields("InvoiceRefNo").Value = Me.InvoiceRefNo.Value
        rs.Fields("InvoiceDate").Value = Me.InvoiceDate.Value
        rs.Fields("DueDate").Value = Me.DueDate.Value
        rs.Fields("TotalCost").Value = Me.TotalCost.Value
        rs.Fields("VAT").Value = Me.VAT.Value
        rs.Fields("TotalGross").Value = Me.TotalGross.Value
        rs.Fields("FuelStatus").Value = Me.FuelStatus.Value
        ' Save the New record
        rs.Update

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Reset the form fields
        InvoiceReset

        MsgBox "Data added successfully."
        Me.Undo
    End If
    Exit Sub

'ErrorHandler:
    'MsgBox "An unexpected error occurred: " & err.Description, vbCritical, "Error"
    'Me.Undo
End Sub

Private Sub ServiceStatus_AfterUpdate()
    Dim CurrentServiceStatus As String
    Dim strSQL As String
    Dim CurrentOperator As String
    Dim CurrentCustomer As String

    CurrentCustomer = Me.CustomerName.Value
    CurrentOperator = Me.OperatorListCb.Value
    CurrentServiceStatus = Me.ServiceStatus.Value

    ' Use a Function To build the query safely
    strSQL = BuildSafeRefNoQueryFormServiceStatus(CurrentServiceStatus)

    ' Set the RowSource of the combo box To the dynamic query
    Me.RefNo.RowSource = strSQL
    Me.RefNo.Requery
End Sub

Private Function BuildSafeRefNoQueryFormServiceStatus(ByVal ServiceStatusValue As String) As String
    Dim CurrentOperator As String
    Dim CurrentCustomer As String

    CurrentCustomer = Me.CustomerName.Value
    CurrentOperator = Me.OperatorListCb.Value

    ' Build the SQL query With And conditions For Customer And Operator
    BuildSafeRefNoQueryFormServiceStatus = "Select DISTINCT RefNo, ID FROM FlightSupport_FuelPreformaQ_New_Test " & _
    "WHERE Customer = '" & Replace(CurrentCustomer, "'", "''") & "' " & _
    "And Operator = '" & Replace(CurrentOperator, "'", "''") & "' " & _
    "And FuelStatus = '" & Replace(ServiceStatusValue, "'", "''") & "'"
End Function
Private Sub EndDate_AfterUpdate()
    Dim strSQL As String
    Dim CurrentStart As Date
    Dim CurrentEnd As Date
    ' **Remove the \# characters from conJetDate**
    Const conJetDate = "dd/mm/yyyy"

    ' Assign values to CurrentStart and CurrentEnd outside the If blocks
    CurrentStart = Me.StartDate.Value
    CurrentEnd = Me.EndDate.Value

    ' Format CurrentStart and CurrentEnd to conJetDate format
    CurrentStart = Format(CurrentStart, conJetDate)
    CurrentEnd = Format(CurrentEnd, conJetDate)

    If Not IsNull(CurrentStart) And Not IsNull(CurrentEnd) Then
        ' Build the SQL query
        strSQL = "SELECT DISTINCT RefNo, ID FROM " & _
                 "(" & _
                 "SELECT RefNo, ID, DepartDate FROM FlightSupport_FuelPreformaQ_New_Test" & _
                 ") AS TempTable " & _
                 "WHERE DepartDate >= #" & CurrentStart & "# " & _
                 "AND DepartDate < #" & CurrentEnd & "#"

        Me.RefNo.RowSource = strSQL
        Me.RefNo.Requery
    Else
        If IsNull(Me.StartDate) Then
            MsgBox "Please add Start date", vbCritical, "Tahseen Pro"
        Else
            MsgBox "Please add End date", vbCritical, "Tahseen Pro"
        End If
        Exit Sub
    End If
End Sub

Private Function BuildSafeRefNoQueryFormDate(ByVal StartDate As Date, ByVal EndDate As Date) As String
    Const conJetDate = "\#dd\/mm\/yyyy\#"
    ' Build the SQL query With only DepartDate conditions
    BuildSafeRefNoQueryFormDate = "Select DISTINCT RefNo, ID FROM FlightSupport_FuelPreformaQ_New_Test " & _
    "WHERE FuelReleaseT_New.DepartDate >= '" & Replace(Format(StartDate, conJetDate), "'", "''") & "' " & _
    "And FuelReleaseT_New.DepartDate < '" & Replace(Format(EndDate + 1, conJetDate), "'", "''") & "'"
End Function
Private Sub EndDate_Click()
'    InputDateField EndDate, "Select a date To use this on your form"
End Sub
Private Sub StartDate_Click()
    InputDateField StartDate, "Select a date To use this on your form"
End Sub

Private Sub resetRowSource()
Dim strSQL As String
strSQL = ""
Me.RefNo.RowSource = strSQL
Me.OperatorListCb.RowSource = strSQL
Me.RefNo_Airport.RowSource = strSQL

End Sub

