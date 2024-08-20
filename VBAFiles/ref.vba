Private Sub RefNo_AfterUpdate()
    Dim selectedRefNo As Variant
    Dim query As String

    selectedRefNo = Me.RefNo.Value ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select Sect_ID FROM NewSectorsT WHERE RefNo = '" & selectedRefNo & "';"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query

    Forms("LASTFORM").Controls("LastForm").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    ' Requery the other subform's combo box To reflect the updated options
    Forms("LASTFORM").Controls("LastForm").Form.Controls("PermitBriefF").Controls("Sect_ID").Requery
End Sub

Private Sub AIRPORT_AfterUpdate()
    Dim selectedAirport As String
    Dim selectedVendorTable As String

    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Value ' Get the selected value from AIRPORT combo box
    query = "Select * FROM " & selectedVendorTable

    ' Assign the control source For each field based on the selected airport value
    Me.BasePrice.ControlSource= "selectedVendorTable." & selectedAirport & ".BasePrice"
    Me.Notes.ControlSource= "selectedVendorTable." & selectedAirport & ".Notes"
    Me.Taxes.ControlSource= "selectedVendorTable." & selectedAirport & ".Taxes"
    Me.EstimatedTotal.ControlSource= "selectedVendorTable." & selectedAirport & ".EstimatedTotal"
    Me.IPA.ControlSource= "selectedVendorTable." & selectedAirport & ".IPA"
    Me.FuelType.ControlSource= "selectedVendorTable." & selectedAirport & ".FuelType"
    Me.FlightType.ControlSource= "selectedVendorTable." & selectedAirport & ".FlightType"
    Me.FBO.ControlSource= "selectedVendorTable." & selectedAirport & ".FBO"
    Me.Validity.ControlSource= "selectedVendorTable." & selectedAirport & ".Validity"

    ' Refresh the form To reflect the updated control sources
    Me.Refresh
End Sub


Private Sub AIRPORT_AfterUpdate()
    Dim selectedAirport As String
    Dim selectedVendorTable As String

    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Value ' Get the selected value from AIRPORT combo box
    query = "Select * FROM " & selectedVendorTable

    ' Assign the control source For each field based on the selected airport value
    Me.BasePrice.ControlSource= query & ".BasePrice"
    Me.Notes.ControlSource= query & ".Notes"


    ' Refresh the form To reflect the updated control sources
    Me.Refresh
End Sub




Private Sub AIRPORT_AfterUpdate()
    Dim selectedAirport As String
    Dim selectedVendorTable As String
    Dim query As String

    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Value ' Get the selected value from AIRPORT combo box
    query = "Select * FROM " & selectedVendorTable

    ' Assign the control source For each field based on the selectedVendorTable And the Airport value
    Me.BasePrice.ControlSource = query & ".BasePrice"
    Me.Notes.ControlSource = query & ".Notes"
    ' Assign control source For other fields here

    ' Refresh the form To reflect the updated control sources
    Me.Requery

End Sub


Private Sub AIRPORT_AfterUpdate()
    Dim selectedAirport As String
    Dim selectedVendorTable As String

    selectedVendorTable = Me.Vendor.Value

    Me.RecordSource = "Select * FROM " & selectedVendorTable & " WHERE ID = '" & Me.AIRPORT.Column(0) & "';"

    Me.Requery
End Sub


Private Sub AIRPORT_AfterUpdate()
    Dim selectedAirport As Integer
    Dim selectedVendorTable As String
    Dim VendorPriceQuery As String
    Dim ValidityQuery As String

    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Column(0)

    If Not IsNull(Me.AIRPORT.Value) Then
        VendorPriceQuery = "Select VendorPrice FROM " & selectedVendorTable & " WHERE ID = '" & selectedAirport & "';"
        me.baseprice.rowsource=VendorPriceQuery
        Me.baseprice.Requery
        ValidityQuery = "Select Validity FROM " & selectedVendorTable & " WHERE ID = '" & selectedAirport & "';"
        me.Validity.rowsource=ValidityQuery
        Me.Validity.Requery
    Else
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
    End If
End Sub











Private Sub Vendor_AfterUpdate()
    Dim selectedVendorTable As String
    Dim query As String

    selectedVendorTable = Me.Vendor.Value ' Get the selected value from Vendor combo box

    RecordQuery = "Select Airport, ID, VendorPrice, Validity, Notes, Taxes, FBO, IPA, EstimatedTotal, FlightType, FuelType FROM " & selectedVendorTable
    Me.RecordSource = RecordQuery

    ' Construct the dynamic query based on the selected vendor table
    query = "Select Airport, ID FROM " & selectedVendorTable

    ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    Me.AIRPORT.RowSource = query

    ' Requery the AIRPORT combo box To reflect the updated options
    Me.AIRPORT.Requery
End Sub

Private Sub Vendor_AfterUpdate()
    Dim selectedVendorTable As String
    Dim query As String
    Dim RecordQuery As String
    selectedVendorTable = Me.Vendor.Value ' Get the selected value from Vendor combo box

    RecordQuery = "Select Airport, ID, VendorPrice, Validity, Notes, Taxes, FBO, IPA, EstimatedTotal, FlightType, FuelType FROM " & selectedVendorTable
    Me.RecordSource = RecordQuery

    ' Construct the dynamic query based on the selected vendor table
    query = "Select Airport, ID FROM " & selectedVendorTable

    ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    Me.AIRPORT.RowSource = query

    ' Requery the AIRPORT combo box To reflect the updated options
    Me.AIRPORT.Requery

    Debug.Print "Record Query: " & RecordQuery
End Sub
///////////////////////////////////////////////////////////////////////

Private Sub Airport_AfterUpdate()
    Dim selectedAirport As As Variant
    Dim LowestPriceQ As String
    selectedAirport = Me.AIRPORT.Column(0)

    If Not IsNull(selectedAirport) Then
        LowestPriceQ = "Select RecordSource, VendorPrice FROM FuelSupport_PriceQ WHERE Airport = " & selectedAirport & " And VendorPrice = (Select MIN(VendorPrice) FROM FuelSupport_PriceQ WHERE Airport = " & selectedAirport & ")"
        Me.LowestPrice.RowSource = LowestPriceQ
    End If
End Sub

Private Sub Airport_AfterUpdate()
    Dim selectedAirport As Variant
    Dim airportCode As String
    Dim LowestPriceQ As String

    selectedAirport = Me.AIRPORT

    If Not IsNull(selectedAirport) Then

        LowestPriceQ = "Select MinVendorPrice FROM FuelSupport_PriceQ WHERE Airport = '" & selectedAirport & "'"

        Me.LowestPriceTXT = LowestPriceQ

    End If
End Sub


Private Sub Airport_AfterUpdate()
    Dim selectedAirport As Variant
    Dim airportCode As String
    Dim LowestPriceQ As String
    Dim minVendorPrice As Variant

    selectedAirport = Me.AIRPORT

    If Not IsNull(selectedAirport) Then
        airportCode = selectedAirport

        LowestPriceQ = "Select MinVendorPrice FROM FuelSupport_PriceQ WHERE Airport = '" & airportCode & "'"

        minVendorPrice = DLookup("MinVendorPrice", "FuelSupport_PriceQ", "Airport = '" & airportCode & "'")

        If Not IsNull(minVendorPrice) Then
            Me.LowestPriceTXT.Value = minVendorPrice
        Else
            Me.LowestPriceTXT.Value = "No data found"
        End If
    End If
End Sub




Private Sub Vendor_AfterUpdate()
    RenderRecord

    ' Dim selectedVendorTable As String
    ' Dim query As String
    ' Dim RecordQuery As String
    ' selectedVendorTable = Me.Vendor.Value ' Get the selected value from Vendor combo box

    ' RecordQuery = "Select Airport, ID, VendorPrice, Validity, Notes, Taxes, FBO, IPA, EstimatedTotal, FlightType, FuelType FROM " & selectedVendorTable
    ' Me.RecordSource = RecordQuery

    ' ' Construct the dynamic query based on the selected vendor table
    ' query = "Select   Airport,ID FROM " & selectedVendorTable

    ' ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    ' Me.AIRPORT.RowSource = query

    ' ' Requery the AIRPORT combo box To reflect the updated options
    ' Me.AIRPORT.Requery

    ' Debug.Print "Record Query: " & RecordQuery
End Sub




Private Sub RenderRecord()
    Dim selectedVendorTable As String
    Dim selectedAirport As Integer
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Dim FlightType As String

    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Column(1)
    FlightType = Me.FlightType.Value

    If Not IsNull(selectedAirport) Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)

        ' Find the row that matches the selectedAirport And FlightType
        rs.FindFirst "ID = " & selectedAirport & " And FlightType = '" & FlightType & "'"

        If Not rs.NoMatch Then
            ' Assign values To the fields in the form
            'Me.AIRPORT.Value = rs("Airport").Value
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.EstimatedTotal.Value = rs("EstimatedTotal").Value
            Me.IPA.Value = rs("IPA").Value
            Me.FuelType.Value = rs("FuelType").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.FBO.Value = rs("FBO").Value
            Me.Validity.Value = rs("Validity").Value

            rs.Close
            Set rs = Nothing
            Set db = Nothing
            'MsgBox "Data rendered successfully."
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    Else
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
    End If
End Sub
////////////////////////////////////////////////////////////
Private Sub AIRPORT_AfterUpdate()
    Dim selectedAirport As Integer
    Dim selectedVendorTable As String
    Dim VendorPriceQuery As String
    Dim ValidityQuery As String

    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Column(1)

    If Not IsNull(selectedAirport) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        'find the row we need To match
        rs.FindFirst "ID = '" & selectedAirport & "'"

        If Not rs.NoMatch Then
            RenderRecord
        Else
            MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
        End If
     Exit Sub
End Sub

Private Sub AIRPORT_AfterUpdate()
    RenderRecord
End Sub

Private Sub RenderRecord()
    Dim selectedVendorTable As String
    Dim selectedAirport As Integer
    Dim rs As DAO.Recordset
    Dim db As DAO.Database

    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Column(1)

    If Not IsNull(selectedAirport) Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)

        ' Find the row that matches the selectedAirport
        rs.FindFirst "ID = " & selectedAirport

        If Not rs.NoMatch Then
            ' Assign values To the fields in the form
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.EstimatedTotal.Value = rs("EstimatedTotal").Value
            Me.IPA.Value = rs("IPA").Value
            Me.FuelType.Value = rs("FuelType").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.FBO.Value = rs("FBO").Value
            Me.Validity.Value = rs("Validity").Value

            ' rs.Close
            ' Set rs = Nothing
            ' Set db = Nothing
            ' me.undo
            MsgBox "Data rendered successfully."
        Else
            MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
        End If
    End If
End Sub
Private Sub UpdateFuelprice_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim selectedVendorTable As String
    Dim selectedAirport As Integer
    'Dim FlightType As String

    'FlightType = Me.FlightType.Value
    selectedVendorTable = Me.Vendor.Value
    selectedAirport = Me.AIRPORT.Column(1)

    Set db = CurrentDb
    Set rs = db.OpenRecordset(selectedVendorTable, dbOpenDynaset)

    If Not rs.EOF Then
        rs.MoveFirst ' Move To the first record in the recordset

        ' Loop through the recordset
        Do Until rs.EOF
            If rs("ID").Value = selectedAirport Then
                rs.Edit
                rs.Fields("VendorPrice").Value = Me.BasePrice.Value
                rs.Fields("Notes").Value = Me.Notes.Value
                rs.Fields("Taxes").Value = Me.Taxes.Value
                rs.Fields("EstimatedTotal").Value = Me.EstimatedTotal.Value
                rs.Fields("IPA").Value = Me.IPA.Value
                rs.Fields("FuelType").Value = Me.FuelType.Value
                rs.Fields("FlightType").Value = Me.FlightType.Value
                rs.Fields("FBO").Value = Me.FBO.Value
                rs.Fields("Validity").Value = Me.Validity.Value
                rs.Update
            End If

            rs.MoveNext ' Move To the Next record
        Loop

        rs.Close
        Set rs = Nothing
        Set db = Nothing

        Me.Requery
        ResetBtn_Click
        MsgBox "Data updated successfully."
    Else
        MsgBox "No records found.", vbExclamation + vbOKOnly, "Update Failed"
    End If
End Sub

Private Sub UpdateHadidFUEL()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("HadidFuelT", dbOpenDynaset)

    ' Assign values To the fields in the table
    Me.BasePrice.Value= rs("VendorPrice").Value
    Me.Notes.Value= rs("Notes").Value
    Me.Taxes.Value= rs("Taxes").Value
    Me.EstimatedTotal.Value= rs("EstimatedTotal").Value
    Me.IPA.Value= rs("IPA").Value
    Me.FuelType.Value= rs("FuelType").Value
    Me.FlightType.Value= rs("FlightType").Value
    Me.FBO.Value= rs("FBO").Value
    Me.AIRPORT.Value= rs("AIRPORT").Value
    Me.Validity.Value= rs("Validity").Value
    Me.Vendor.Value= rs("Vendor").Value

    ' Save the New record
    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset the form fields
    FuelReset

    MsgBox "Data Updated successfully."
    Me.Undo
End Sub




/////////////////////////////////



Private Sub AddFuelQuote_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("ASMFuelT", dbOpenDynaset)

    If IsNull(Me.BasePrice.Value) Or IsNull(Me.Validity.Value) Or IsNull(Me.OPSDate.Value) Or IsNull(Me.Customer.Value) Or IsNull(Me.Operator.Value) _
        Or IsNull(Me.ACType.Value) Or IsNull(Me.ACMTOW.Value) Or IsNull(Me.ACRegCb.Value) Or IsNull(Me.PFL.Value) Or IsNull(Me.Schedule.Value) Then
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
        rs.Fields("TASPrice").Value = Me.Client.Value
        rs.Fields("QuoteRefNo").Value = Me.QuoteRefNo.Value
        rs.Fields("Vendor").Value = Me.vendorName.Value
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

End Sub

Private Sub Vendorname()
    Dim vendorName As String
    Dim vendorTable As String
    vendorName = me.vendorName.value

    vendorTable=Me.Vendor.Value
    If Not isnull(vendorTable) Then 
        If (vendorTable = "ASMFuelT") Then 
            vendorName= "ASM"
        Elseif(vendorTable = "HadidFuelT")
            vendorName= "Hadid"
        Elseif(vendorTable = "LINKAEROFuelT")
            vendorName= "LINKAERO"
        Elseif(vendorTable = "UCIGFuelT")
            vendorName= "UCIG"
        Elseif(vendorTable = "AEGFuelT")
            vendorName= "AEG"
        End If

    End If
end Sub