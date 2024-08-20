Option Compare Database

Private Sub Airport_AfterUpdate()
    MinPrice_MinVendor
End Sub
Private Sub MinPrice_MinVendor()
    Dim selectedAirport As Variant
    Dim airportCode As String
    Dim LowestPriceQ As String
    Dim minVendorPrice As Variant
    Dim vendorTable As String
    Dim AirportIDQ As Integer

    selectedAirport = Me.AIRPORT

    If Not IsNull(selectedAirport) Then
        airportCode = selectedAirport

        LowestPriceQ = "Select MinVendorPrice, TableName,ID FROM FuelSupport_PriceQ WHERE Airport = '" & airportCode & "'"

        minVendorPrice = DLookup("MinVendorPrice", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
        vendorTable = DLookup("TableName", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
        AirportIDQ = DLookup("ID", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")

        If Not IsNull(minVendorPrice) Then
            Me.AirportID.Value = AirportIDQ
            Me.LowestVendor.Value = vendorTable
            Me.Vendor.Value = vendorTable
            Me.LowestPriceTxt.Value = minVendorPrice
            RenderRecord
        Else
            Me.LowestPriceTxt.Value = "No data found"
        End If
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
    AirportID = Me.AirportID
    selectedAirport = Me.AIRPORT.Value
    vendorNameSub
    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        'If Not IsNull(selectedAirport) Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)

        ' Find the row that matches the selectedAirport And FlightType
        'rs.FindFirst "ID = " & AirportID & " And FlightType = '" & FlightType & "'"
        rs.FindFirst "ID = " & AirportID
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form
            'Me.AIRPORT.Value = rs("Airport").Value
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
    On Error Goto ErrorHandler

        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("ASMFuelT", dbOpenDynaset)

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
            rs.Fields("IPA").Value = Me.IPA.Value
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
            ' Save the New record
            rs.Update

            ' Clean up
            rs.Close
            Set rs = Nothing
            Set db = Nothing

            ' Reset the form fields
            FuelReset

            MsgBox "Data added successfully.", vbInformation, "Success"
            Me.Undo
        End If

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & err.Description, vbCritical, "Error"
        rs.Close
        Set rs = Nothing
        Set db = Nothing
End Sub
Private Sub vendorNameSub()
    Dim vendorName As String
    Dim vendorTable As String
    vendorName = Me.vendorName.Value

    vendorTable = Me.Vendor.Value
    If Not IsNull(vendorTable) Then
        If (vendorTable = "ASMFuelT") Then
            vendorName = "ASM"
        Elseif (vendorTable = "HadidFuelT") Then
            vendorName = "Hadid"
        Elseif (vendorTable = "LINKAEROFuelT") Then
            vendorName = "LINKAERO"
        Elseif (vendorTable = "UCIGFuelT") Then
            vendorName = "UCIG"
        Elseif (vendorTable = "AEGFuelT") Then
            vendorName = "AEG"
        End If

    End If
End Sub

Private Sub cmdGenerate_Click()
    Dim RefValue As Integer
    Dim rst As DAO.Recordset

    RefValue = Nz(DLookup("Quote_ID", "TRefNo"), 0)

    Me.QuoteRefNo = "TAS_FQ-" & Format(RefValue + 1, "000") & "/" & Format(Date, "yy")


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
Private Sub Vendor_Change()
    Dim selectedVendorTable As String
    Dim query As String
    Dim RecordQuery As String
    selectedVendorTable = Me.Vendor.Value ' Get the selected value from Vendor combo box

    RecordQuery = "Select Airport, ID, VendorPrice, Validity, Notes, Taxes, FBO, IPA, EstimatedTotal, FlightType, FuelType FROM " & selectedVendorTable
    Me.RecordSource = RecordQuery

    ' Construct the dynamic query based on the selected vendor table
    query = "Select   Airport,ID FROM " & selectedVendorTable

    ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    Me.AIRPORT.RowSource = query

    ' Requery the AIRPORT combo box To reflect the updated options
    Me.AIRPORT.Requery

    Debug.Print "Record Query: " & RecordQuery
End Sub
Private Sub Form_Load()
    Dim fw As New clFormWindow

    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
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
    'Me.Vendor.Value = Null
    Me.Undo
End Sub
Private Sub Validity_Click()
    InputDateField Validity, "Select a date To use this on your form"
End Sub
Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FuelSupport_FuelQuote", acSaveNo
End Sub


////////////////////////////////////////



Select T1.Airport, T1.MinVendorPrice, T2.TableName, T2.ID
FROM 
(
Select Airport, MIN(VendorPrice) As MinVendorPrice 
FROM 
(
Select Airport, VendorPrice, 'AEGFuelT' As TableName FROM AEGFuelT
UNION ALL
Select Airport, VendorPrice, 'ASMFuelT' As TableName FROM ASMFuelT
UNION ALL
Select Airport, VendorPrice, 'UCIGFuelT' As TableName FROM UCIGFuelT
UNION ALL
Select Airport, VendorPrice, 'LINKAEROFuelT' As TableName FROM LINKAEROFuelT
UNION ALL
Select Airport, VendorPrice, 'HadidFuelT' As TableName FROM HadidFuelT
) As AllVendorPrices 
GROUP BY Airport
HAVING COUNT(*) > 1
) As T1
INNER JOIN 
(
Select ID, Airport, VendorPrice, TableName 
FROM 
(
Select ID, Airport, VendorPrice, 'AEGFuelT' As TableName FROM AEGFuelT
UNION ALL
Select ID, Airport, VendorPrice, 'ASMFuelT' As TableName FROM ASMFuelT
UNION ALL
Select ID, Airport, VendorPrice, 'UCIGFuelT' As TableName FROM UCIGFuelT
UNION ALL
Select ID, Airport, VendorPrice, 'LINKAEROFuelT' As TableName FROM LINKAEROFuelT
UNION ALL
Select ID, Airport, VendorPrice, 'HadidFuelT' As TableName FROM HadidFuelT
) As AllVendorPrices
) As T2 
ON T1.Airport = T2.Airport 
And T1.MinVendorPrice = T2.VendorPrice;




Dim selectedAirport As Variant
Dim airportCode As String
Dim LowestPriceQ As String
Dim minVendorPrice As Variant
Dim vendorTable As String
Dim AirportIDQ As Integer

selectedAirport = Me.AIRPORT

If Not IsNull(selectedAirport) Then
    airportCode = selectedAirport

    LowestPriceQ = "Select MinVendorPrice, TableName,ID FROM FuelSupport_PriceQ WHERE Airport = '" & airportCode & "'"

    minVendorPrice = DLookup("MinVendorPrice", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
    vendorTable = DLookup("TableName", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")
    AirportIDQ = DLookup("ID", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")

    If Not IsNull(minVendorPrice) Then
        Me.AirportID.Value = AirportIDQ
        Me.LowestVendor.Value = vendorTable
        Me.Vendor.Value = vendorTable
        Me.LowestPriceTxt.Value = minVendorPrice
        RenderRecord
    Else
        Me.LowestPriceTxt.Value = "No data found"
    End If
End If
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

Private Sub ShowAllVendorsPrice()
    Dim ICAO As Variant
    ICAO = Me.VendorLocCmb.Value
    selectedColumn = Me.FuelVendorsCmb.Value ' Get the selected column name from the combo box

    If Not IsNull(selectedColumn) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("FuelSupplierT", dbOpenSnapshot)

        rs.FindFirst "ICAO = '" & ICAO & "'"

        If Not rs.NoMatch Then
            Dim columnName As DAO.Field
            Set columnName = rs.Fields(selectedColumn)

            If Not columnName Is Nothing Then
                Me.VendorPrice.Value = columnName.Value
                Me.Validity.Value = rs("Validity").Value
            Else
                Me.VendorPrice.Value = Null
                Me.Validity.Value = Null
            End If

        Else
            Me.VendorPrice.Value = Null
            Me.Validity.Value = Null

        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.VendorPrice.Value = Null
        Me.Validity.Value = Null

    End If

 Exit Sub
End Sub