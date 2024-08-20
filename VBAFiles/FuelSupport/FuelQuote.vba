Option Compare Database

Private Sub Airport_AfterUpdate()
    MinPrice_MinVendor
End Sub
Private Sub MinPrice_MinVendor()
    Dim selectedAirport As Variant
    Dim airportCode As String
    Dim LowestPriceQ As String
    Dim MinVendorPrice As Variant
    Dim vendorTable As String
    Dim AirportIDQ As Integer

    selectedAirport = Me.AIRPORT

    If Not IsNull(selectedAirport) Then
        airportCode = selectedAirport

        LowestPriceQ = "Select MinVendorPrice, TableName,ID FROM FuelSupport_PriceQ WHERE Airport = '" & airportCode & "'"

        MinVendorPrice = DLookup("MinVendorPrice", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
        vendorTable = DLookup("TableName", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
        AirportIDQ = DLookup("ID", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")

        If Not IsNull(MinVendorPrice) Then
            Me.AirportID.Value = AirportIDQ
            Me.LowestVendor.Value = vendorTable
            Me.Vendor.Value = vendorTable
            Me.LowestPriceTxt.Value = MinVendorPrice
            RenderRecord
        Else
            MsgBox "No data found", vbCritical, "My Flight App"

        End If
    End If
End Sub

Private Sub Command35_Click()
    DoCmd.OpenForm "FuelSupport_AddFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command36_Click()
    DoCmd.OpenForm "FuelSupport_updateFuel", WindowMode:=acWindowNormal
End Sub

Private Sub PrintQuote_Click()
    If Not IsNull(Me.QuoteRefNo) Then

        Dim strWehre As String
        Dim InLeng As Long

        Const conJetDate = "\#dd\/mm\/yyyy\#"

        If Not IsNull(Me.Client) Then

            strWhere = strWhere & "([Client] Like ""*" & Me.Client & "*"") And "
        End If

        If Not IsNull(Me.QuoteRefNo) Then

            strWhere = strWhere & "([QuoteRefNo] Like ""*" & Me.QuoteRefNo & "*"") And "
        End If

        IngLen = Len(strWhere) - 5
        If IngLen <= 0 Then
            MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
        Else
            strWhere = Left$(strWhere, IngLen)
            Me.Filter = strWhe
            Me.FilterOn = True
            DoCmd.OpenReport "FuelSupport_ExternalReport", acViewPreview, , strWhere
        End If
    Else
        DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
    End If


End Sub

Private Sub Command659_Click()
    Me.Refresh
End Sub


Private Sub RenderRecord()
    Dim selectedVendorTable As String
    Dim selectedAirport As Variant
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    selectedVendorTable = Me.Vendor
    selectedAirport = Me.AIRPORT



    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        ' Find the row that matches the selectedAirport
        rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.EstimatedTotal.Value = rs("EstimatedTotal").Value
            Me.IPA.Value = rs("IPA").Value
            Me.FuelType.Value = rs("FuelType").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.validity.Value = rs("Validity").Value
            Me.VendorCmbPrice.Value = rs("VendorPrice").Value
            Me.Currency.Value = rs("Currency").Value
            Me.Unit.Value = rs("Unit").Value
            rs.Close
            Set rs = Nothing
            Set db = Nothing

            If Me.validity < date Then
                Me.validity.BackColor = RGB(255, 0, 0) ' Set red background color
                Me.validity.ForeColor = RGB(255, 255, 255) ' Set white font color
            Else
                Me.validity.BackColor = RGB(255, 255, 255) ' Set white font color
                Me.validity.ForeColor = RGB(47, 54, 153)  ' Set tas font color
            End If
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    End If
End Sub

Private Sub cmdGenerate_Click()
    Dim RefValue As Integer
    Dim rst As DAO.Recordset
    RefValue = Nz(DLookup("Quote_ID", "TRefNo"), 0)

    Me.QuoteRefNo = "TAS_FQ-" & Format(RefValue + 7, "0000") & "/" & Format(date, "yy")
    Set rst = CurrentDb.OpenRecordset("TRefNo", dbOpenDynaset, dbSeeChanges)

    With rst
        .Edit
        ![Quote_ID] = ![Quote_ID] + 7
        .Update
    End With

    rst.Close
    Set rst = Nothing

End Sub
Private Sub newQuoteBtn_Click()
    If IsNull(Me.QuoteRefNo) Or Me.QuoteRefNo = "" Then
        cmdGenerate_Click
    End If
    ' Set focus To the first field For data entry
    Me.AIRPORT.SetFocus
End Sub

Private Sub Vendor_AfterUpdate()
    RenderRecord
End Sub

Private Sub Form_Load()
    RowSourceReset
End Sub

Private Sub QuoteReset()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    Me.IPA.Value = Null
    Me.AIRPORT.Value = Null

    RowSourceReset
End Sub
Private Sub RowSourceReset()
    Dim AirportQuery As String

    AirportQuery = "Select Airport FROM FuelSupport_PriceQ "

    Me.RecordSource = AirportQuery

    ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    Me.AIRPORT.RowSource = AirportQuery

    ' Requery the AIRPORT combo box To reflect the updated options
    Me.AIRPORT.Requery
End Sub

Private Sub ResetBtn_Click()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    Me.IPA.Value = Null
    Me.FlightType.Value = Null
    Me.AIRPORT.Value = Null
    Me.QuoteRefNo.Value = Null
    RowSourceReset

End Sub

Private Sub Validity_Click()
    InputDateField validity, "Select a date To use this on your form"
End Sub
Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FuelSupport_FuelQuote", acSaveNo
End Sub
Private Sub AddFuelQuote_Click()
    If IsNull(Me.BasePrice.Value) Or IsNull(Me.validity.Value) Or IsNull(Me.Client.Value) Or IsNull(Me.AIRPORT.Value) Or IsNull(Me.QuoteRefNo.Value) _
        Or IsNull(Me.TASPrice.Value) Or IsNull(Me.AddedRate.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    End If

    If CheckIFQuoteIsAlreadyAdded() Then
        MsgBox "Quote For the same airport already exists."
     Exit Sub
    End If

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("FuelSupport_FuelQuotationsT", dbOpenDynaset)

    rs.AddNew
    rs.Fields("VendorPrice").Value = Me.BasePrice.Value
    rs.Fields("Notes").Value = Me.Notes.Value
    rs.Fields("Taxes").Value = Me.Taxes.Value
    rs.Fields("EstimatedTotal").Value = Me.EstimatedTotal.Value
    rs.Fields("TASEstimatedTotal").Value = Me.TASEstimatedTotal.Value
    rs.Fields("IPA").Value = Me.IPA.Value
    rs.Fields("FuelType").Value = Me.FuelType.Value
    rs.Fields("FlightType").Value = Me.FlightType.Value
    rs.Fields("Airport").Value = Me.AIRPORT.Value
    rs.Fields("Currency").Value = Me.Currency.Value
    rs.Fields("Unit").Value = Me.Unit.Value
    rs.Fields("EffDate").Value = Me.EffDate.Value
    rs.Fields("Client").Value = Me.Client.Value
    rs.Fields("TASPrice").Value = Me.TASPrice.Value
    rs.Fields("QuoteRefNo").Value = Me.QuoteRefNo.Value
    rs.Fields("AddedRate").Value = Me.AddedRate.Value
    rs.Fields("Vendor").Value = Me.vendorName.Value ' Add this line
    rs.Fields("Validity").Value = Me.validity.Value
    ' Save the New record
    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset the form fields
    QuoteReset

    MsgBox "Data added successfully.", vbInformation, "Success"
    Me.Undo
End Sub

Private Function CheckIFQuoteIsAlreadyAdded() As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim currentQuoteRefNo As Variant
    Dim selectedAirport As Variant

    currentQuoteRefNo = Me.QuoteRefNo.Value
    selectedAirport = Me.AIRPORT.Value

    Set db = CurrentDb
    Set rs = db.OpenRecordset("FuelSupport_FuelQuotationsT", dbOpenDynaset)

    If Not rs.EOF Then
        rs.MoveFirst ' Move To the first record in the recordset

        ' Loop through the recordset
        Do Until rs.EOF
            If rs("QuoteRefNo").Value = currentQuoteRefNo And rs("Airport").Value = selectedAirport Then
                rs.Close
                Set rs = Nothing
                Set db = Nothing
                CheckIFQuoteIsAlreadyAdded = True ' Return True If matching quote is found
             Exit Function ' Exit the Function If a matching quote is found
            End If

            rs.MoveNext ' Move To the Next record
        Loop

        rs.Close
        Set rs = Nothing
        Set db = Nothing

        Me.Requery
    End If

    CheckIFQuoteIsAlreadyAdded = False ' Return False If no matching quote is found
End Function





/////////////////////////////////////////////
old code:


Option Compare Database

Private Sub Airport_AfterUpdate()
    MinPrice_MinVendor
End Sub
Private Sub MinPrice_MinVendor()
    Dim selectedAirport As Variant
    Dim airportCode As String
    Dim LowestPriceQ As String
    Dim MinVendorPrice As Variant
    Dim vendorTable As String
    Dim AirportIDQ As Integer

    selectedAirport = Me.AIRPORT

    If Not IsNull(selectedAirport) Then
        airportCode = selectedAirport

        LowestPriceQ = "Select MinVendorPrice, TableName,ID FROM FuelSupport_PriceQ WHERE Airport = '" & airportCode & "'"

        MinVendorPrice = DLookup("MinVendorPrice", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
        vendorTable = DLookup("TableName", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
        AirportIDQ = DLookup("ID", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")

        If Not IsNull(MinVendorPrice) Then
            Me.AirportID.Value = AirportIDQ
            Me.LowestVendor.Value = vendorTable
            Me.Vendor.Value = vendorTable
            Me.LowestPriceTxt.Value = MinVendorPrice
            RenderRecord
        Else
            Me.LowestPriceTxt.Value = "No data found"
        End If
    End If
End Sub



Private Sub Command35_Click()
    DoCmd.OpenForm "FuelSupport_AddFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command36_Click()
    DoCmd.OpenForm "FuelSupport_updateFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command42_Click()
    DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
End Sub

Private Sub PrintQuote_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#dd\/mm\/yyyy\#"


    If Not IsNull(Me.Client) Then

        strWhere = strWhere & "([Client] Like ""*" & Me.Client & "*"") And "
    End If

    If Not IsNull(Me.QuoteRefNo) Then

        strWhere = strWhere & "([QuoteRefNo] Like ""*" & Me.QuoteRefNo & "*"") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhe
        Me.FilterOn = True
        DoCmd.OpenReport "FuelSupport_ExternalReport", acViewPreview, , strWhere
    End If

End Sub



Private Sub Command659_Click()
    Me.Refresh
End Sub


Private Sub RenderRecord()
    Dim selectedVendorTable As String
    Dim selectedAirport As Variant
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    selectedVendorTable = Me.Vendor
    'AirportID = Me.AirportID
    selectedAirport = Me.AIRPORT


    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        'If Not IsNull(selectedAirport) Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)

        ' Find the row that matches the selectedAirport And FlightType
        'rs.FindFirst "ID = " & AirportID & " And FlightType = '" & FlightType & "'"
        rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.EstimatedTotal.Value = rs("EstimatedTotal").Value
            'Me.IPA.Value = rs("IPA").Value
            Me.FuelType.Value = rs("FuelType").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.FBO.Value = rs("FBO").Value
            Me.Validity.Value = rs("Validity").Value
            Me.VendorCmbPrice.Value = rs("VendorPrice").Value
            Me.Currency.Value = rs("Currency").Value
            Me.Unit.Value = rs("Unit").Value
            rs.Close
            Set rs = Nothing
            Set db = Nothing
            'MsgBox "Data rendered successfully."
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If



    End If
End Sub
Private Sub AddFuelQuote_Click()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("FuelSupport_FuelQuotationsT", dbOpenDynaset)

    If IsNull(Me.BasePrice.Value) Or IsNull(Me.Validity.Value) Or IsNull(Me.Client.Value) Or IsNull(Me.AIRPORT.Value) Or IsNull(Me.QuoteRefNo.Value) _
        Or IsNull(Me.TASPrice.Value) Or IsNull(Me.AddedRate.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    Else
        rs.AddNew
        rs.Fields("VendorPrice").Value = Me.BasePrice.Value
        rs.Fields("Notes").Value = Me.Notes.Value
        rs.Fields("Taxes").Value = Me.Taxes.Value
        rs.Fields("EstimatedTotal").Value = Me.EstimatedTotal.Value
        rs.Fields("TASEstimatedTotal").Value = Me.TASEstimatedTotal.Value
        'rs.Fields("IPA").Value = Me.IPA.Value
        rs.Fields("FuelType").Value = Me.FuelType.Value
        rs.Fields("FlightType").Value = Me.FlightType.Value
        rs.Fields("FBO").Value = Me.FBO.Value
        rs.Fields("Airport").Value = Me.AIRPORT.Value
        rs.Fields("Currency").Value = Me.Currency.Value
        rs.Fields("Unit").Value = Me.Unit.Value
        rs.Fields("EffDate").Value = Me.EffDate.Value
        rs.Fields("Client").Value = Me.Client.Value
        rs.Fields("TASPrice").Value = Me.TASPrice.Value
        rs.Fields("QuoteRefNo").Value = Me.QuoteRefNo.Value
        rs.Fields("AddedRate").Value = Me.AddedRate.Value
        rs.Fields("Vendor").Value = Me.vendorName.Value
        rs.Fields("Validity").Value = Me.Validity.Value
        ' Save the New record
        rs.Update

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Reset the form fields
        QuoteReset

        MsgBox "Data added successfully.", vbInformation, "Success"
        Me.Undo
    End If

 Exit Sub

End Sub


Private Sub cmdGenerate_Click()
    Dim RefValue As Integer
    Dim rst As DAO.Recordset

    RefValue = Nz(DLookup("Quote_ID", "TRefNo"), 0)

    Me.QuoteRefNo = "TAS_FQ-" & Format(RefValue + 1, "000") & "/" & Format(date, "yy")


    Set rst = CurrentDb.OpenRecordset("TRefNo", dbOpenDynaset, dbSeeChanges)

    With rst
        .Edit
        ![Quote_ID] = ![Quote_ID] + 1
        .Update
    End With

    rst.Close
    Set rst = Nothing

End Sub
Private Sub newQuoteBtn_Click()

    If IsNull(Me.QuoteRefNo) Or Me.QuoteRefNo = "" Then
        cmdGenerate_Click
    End If
    ' Set focus To the first field For data entry
    Me.AIRPORT.SetFocus

End Sub

Private Sub Vendor_AfterUpdate()
    'Vendor_Change
    RenderRecord
End Sub

Private Sub Form_Load()
    RowSourceReset
End Sub
Private Sub QuoteReset()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    'Me.EstimatedTotal.Value = Null
    'Me.IPA.Value = Null
    Me.FBO.Value = Null
    'Me.FlightType.Value = Null
    Me.FBO.Value = Null
    Me.AIRPORT.Value = Null

    RowSourceReset
End Sub
Private Sub RowSourceReset()

    Dim AirportQuery As String


    AirportQuery = "Select Airport FROM FuelSupport_PriceQ "

    Me.RecordSource = AirportQuery

    ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    Me.AIRPORT.RowSource = AirportQuery

    ' Requery the AIRPORT combo box To reflect the updated options
    Me.AIRPORT.Requery
End Sub

Private Sub ResetBtn_Click()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    'Me.EstimatedTotal.Value = Null
    Me.FBO.Value = Null
    Me.FlightType.Value = Null
    Me.FBO.Value = Null
    Me.AIRPORT.Value = Null
    'Me.TASPrice.Value = Null
    Me.QuoteRefNo.Value = Null
    RowSourceReset
    'Me.Undo
End Sub
Private Sub Validity_Click()
    InputDateField Validity, "Select a date To use this on your form"
End Sub
Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FuelSupport_FuelQuote", acSaveNo
End Sub




