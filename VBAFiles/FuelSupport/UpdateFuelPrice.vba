Option Compare Database

Private Sub Form_Load()
    Dim fw As New clFormWindow

    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
End Sub

Private Sub AddFuelService_Click()
    Dim selectedVendorTable As Variant
    selectedVendorTable = Me.Vendor.Value ' Get the selected column name from the combo box

    If Not IsNull(selectedVendorTable) Then
        AddFuel
    Else
        MsgBox "Data Not added."
     Exit Sub
    End If
End Sub

Private Sub AddFuel()
    On Error Goto ErrorHandler
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim selectedVendorTable As String
        Dim selectedAirport As String
        'Dim FlightType As String

        'FlightType = Me.FlightType.Value
        selectedVendorTable = Me.Vendor.Value
        selectedAirport = Me.AIRPORT.Value

        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenDynaset)

        rs.AddNew
        ' Assign values To the fields in the table
        rs.Fields("VendorPrice").Value = Me.BasePrice.Value
        rs.Fields("Notes").Value = Me.Notes.Value
        rs.Fields("Taxes").Value = Me.Taxes.Value
        rs.Fields("EstimatedTotal").Value = Me.EstimatedTotal.Value
        rs.Fields("IPA").Value = Me.IPA.Value
        rs.Fields("FuelType").Value = Me.FuelType.Value
        rs.Fields("FlightType").Value = Me.FlightType.Value
        rs.Fields("FBO").Value = Me.FBO.Value
        rs.Fields("Airport").Value = Me.AIRPORT.Value
        rs.Fields("Validity").Value = Me.Validity.Value

        ' Save the New record
        rs.Update

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Reset the form fields
        FuelReset

        MsgBox "Data added successfully."
        Me.Undo
     Exit Sub

 ErrorHandler:
        MsgBox "An unexpected error occurred: " & err.Description, vbCritical, "Error"
        Me.Undo
     Exit Sub

End Sub
Private Sub FuelReset()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    Me.EstimatedTotal.Value = Null
    Me.IPA.Value = Null
    Me.FBO.Value = Null
End Sub


Private Sub ResetBtn_Click()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    Me.EstimatedTotal.Value = Null
    Me.IPA.Value = Null
    Me.FuelType.Value = Null
    Me.FlightType.Value = Null
    Me.FBO.Value = Null
    Me.AIRPORT.Value = Null
    Me.Validity.Value = Null
    Me.Vendor.Value = Null
    Me.Undo
End Sub

Private Sub Validity_Click()
    InputDateField Validity, "Select a date To use this on your form"
End Sub
Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FuelSupport_AddFuel", acSaveNo
End Sub


Private Sub Vendor_AfterUpdate()
    Dim selectedVendorTable As String
    selectedVendorTable = Me.Vendor.Value

    Dim vendorTable As String
    vendorTable = "Select DISTINCT IPA FROM " & selectedVendorTable & " WHERE IPA IS Not NULL And IPA <> ''"

    Me.IPA.RowSource = vendorTable
End Sub
