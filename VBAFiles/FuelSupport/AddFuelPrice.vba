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



/////////////////////////////////

old code

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
    Dim selectedColumn As Variant
    selectedColumn = Me.Vendor.Value ' Get the selected column name from the combo box

    If Not IsNull(selectedColumn) Then
        ' If (selectedColumn = "HADID") Then
        '     HADIDFUEL
        ' Elseif (selectedColumn = "ASM") Then
        '     ASMFUEL
        ' Elseif (selectedColumn = "UCIG") Then
        '     UCIGFUEL
        ' Elseif (selectedColumn = "AEG FUEL") Then

        '     AEGFUEL
        ' Elseif (selectedColumn = "LINK AERO") Then

        '     LINKFUEL
        ' End If

        AddFueltest
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
        Dim selectedFlightType As String

        selectedVendorTable = Me.Vendor.Value
        selectedAirport = Me.AIRPORT.Value
        selectedFlightType = Me.FlightType.Value

        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenDynaset)

        ' Check If the record already exists
        rs.FindFirst "AIRPORT = '" & selectedAirport & "' And FlightType = '" & selectedFlightType & "'"
        If Not rs.NoMatch Then
            MsgBox "Record already exists."
         Exit Sub
        End If

        If IsNull(Me.BasePrice) Or IsNull(Me.FuelType) Or IsNull(Me.FlightType) Or IsNull(Me.Currency) Or IsNull(Me.Unit) Or IsNull(Me.EstimatedTotal) Or IsNull(Me.Validity) Then
            MsgBox "Add missing fields " & err.Description, vbCritical, "Error"
         Exit Sub
        Else
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
            rs.Fields("InternalNotes").Value = Me.InternalNotes.Value
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
        End If
     Exit Sub

 ErrorHandler:
        MsgBox "An unexpected error occurred: " & err.Description, vbCritical, "Error"
        Me.Undo
End Sub

'old

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
        If IsNull(Me.BasePrice) Or IsNull(Me.FuelType) Or IsNull(Me.FuelType) Or IsNull(Me.FlightType) Or IsNull(Me.Currency) Or IsNull(Me.Unit) Or IsNull(Me.EstimatedTotal) Or IsNull(Me.Validity) Then
            MsgBox "Add missing fields " & err.Description, vbCritical, "Error"
         Exit Sub
        Else

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
            rs.Fields("InternalNotes").Value = Me.InternalNotes.Value
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
        End If
End Sub

Private Sub HADIDFUEL()
    On Error Goto ErrorHandler
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("HadidFuelT", dbOpenDynaset)

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
        MsgBox "An error occurred While adding data: " & err.Description, vbCritical, "Error"
     Exit Sub
End Sub


Private Sub UCIGFUEL()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("UCIGFuelT", dbOpenDynaset)

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
End Sub

Private Sub AEGFUEL()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("AEGFuelT", dbOpenDynaset)

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
End Sub

Private Sub LINKFUEL()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("LINKAEROFuelT", dbOpenDynaset)

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
    rs.Fields("Currency").Value = Me.Currency.Value
    rs.Fields("Unit").Value = Me.Unit.Value
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

