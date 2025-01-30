Option Compare Database
Option Explicit

Private Sub AddQouteBtn_Click()
    DoCmd.OpenForm "FuelSupport_FuelQuote", WindowMode:=acWindowNormal
End Sub

Private Sub Form_GotFocus()
    Form_Load
End Sub
Private Sub AddUserName()
    Dim AddUser As Boolean
    Dim UserID As Integer

    UserID = Forms!loggedInUser!txtUserID

    AddUser = DLookup("CreateUser", "UsersT", "ID =" & UserID & "")

    Me.lblUser.Value = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
Private Sub Form_Load()
    ' Call the function to add user name
    AddUserName   
    Me.Qouted_Vendor.Visible = False
    Me.Qouted_Vendorlb.Visible = False
    Me.AddQouteBtn.Visible = False
End Sub

Private Sub Qouted_Vendor_DblClick(Cancel As Integer)
    PrintQuote_Click
End Sub

Private Sub PrintQuote_Click()
    Dim strWhere As String
    Dim InLeng As Long
    Dim qouteRefno As Variant

    qouteRefno = Me.Qouted_Vendor.Column(1)

    Const conJetDate = "\#dd\/mm\/yyyy\#"

    If Not IsNull(qouteRefno) Then
        If Not IsNull(qouteRefno) Then
            strWhere = strWhere & "([Quote_RefNo] Like '*" & qouteRefno & "*') And "
        End If

        InLeng = Len(strWhere) - 5
        If InLeng <= 0 Then
            MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
        Else
            strWhere = Left$(strWhere, InLeng)
            Me.Filter = strWhere
            Me.FilterOn = True
            DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
        End If
    Else
        DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
    End If
End Sub

Private Sub QuoteSentCb_AfterUpdate()
    If IsNull(Me.QuoteSentCb) Then
     Exit Sub
    Else
        If Me.QuoteSentCb.Value = True Then
            Me.Qouted_Vendor.Visible = True
            Me.Qouted_Vendorlb.Visible = True
            Me.AddQouteBtn.Visible = True
            LookupQoutedVendors
        Elseif Me.QuoteSentCb.Value = False Then
            Me.AddQouteBtn.Visible = False
            Me.Qouted_Vendor.Visible = False
            Me.Qouted_Vendorlb.Visible = False
        End If
    End If
End Sub

Private Sub LookupQoutedVendors()
    Dim selectedAirport As String
    Dim QoutedVendorsQ As String
    Dim Client As String
    Dim strCurrentDate As String

    ' Store current date in correct Access date literal format
    strCurrentDate = Format(date, "dd\/mm\/yyyy")

    ' Retrieve values
    Client = me.CustomerTxt.value
    selectedAirport = Nz(Me.AIRPORT.value, "")

    ' Build query
    QoutedVendorsQ = "Select Vendor, Quote_RefNo FROM FuelSupport_FuelQuotationsT " & _
    "WHERE Airport = '" & selectedAirport & "' " & _
    "ORDER BY ID DESC;"

    ' Set RowSource And requery
    Me.Qouted_Vendor.RowSource = QoutedVendorsQ
    Me.Qouted_Vendor.Requery
End Sub

Private Sub Airport_AfterUpdate()
    MinPrice_MinVendor
End Sub

Private Sub MinPrice_MinVendor()
    Dim selectedAirport As Variant
    'Dim airportCode As String
    Dim LowestPriceQ As String
    Dim MinVendorPrice As Variant
    Dim AirportIDQ As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    selectedAirport = Me.AIRPORT.value

    If Not IsNull(selectedAirport) Then
        ' airportCode = selectedAirport
        Set db = CurrentDb
        Set rs = db.OpenRecordset("Select * FROM FuelSupport_NewFuelQ WHERE Airport = '" & selectedAirport & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            Me.AirportID.Value = rs!ID
            Me.LowestVendor.Value = rs!TableName
            Me.Vendor.Value = rs!TableName
            Me.LowestPriceTxt.Value = rs!MinPrice
            Me.LowCurrency.Value = rs!LowCurrency

            RenderRecord
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    End If
End Sub
Private Sub MinPrice_MinVendor_FlightType()
    Dim selectedAirport As Variant
    Dim LowestPriceQ As String
    Dim selectedFlightType As String
    Dim MinVendorPrice As Variant
    Dim AirportIDQ As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    selectedAirport = Me.AIRPORT.Column(0)
    selectedFlightType = Me.FlightType
    If Not IsNull(selectedAirport) Then
        ' airportCode = selectedAirport
        Set db = CurrentDb
        Set rs = db.OpenRecordset("Select * FROM FuelSupport_NewFuelQ WHERE FlightType = '" & selectedFlightType & "' And Airport = '" & selectedAirport & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            Me.AirportID.Value = rs!ID
            Me.LowestVendor.Value = rs!TableName
            Me.Vendor.Value = rs!TableName
            Me.LowestPriceTxt.Value = rs!MinPrice

            RenderRecord_FlightType
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    End If
End Sub
Private Sub FlightType_AfterUpdate()
    MinPrice_MinVendor_FlightType
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
    selectedAirport = Me.AIRPORT.Column(0)



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
            Me.EstimatedTotaltxt.Value = rs("EstimatedTotal").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.Validity.Value = rs("Validity").Value
            Me.Currency.Value = rs("Currency").Value
            ' will be hidden As no need To show them
            Me.FuelType.Value = rs("FuelType").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.Unit.Value = rs("Unit").Value
            Me.InternalNotes.Value = rs("InternalNotes").Value

            rs.Close
            Set rs = Nothing
            Set db = Nothing

            If Me.Validity < date Then
                Me.Validity.BackColor = RGB(255, 0, 0) ' Set red background color
                Me.Validity.ForeColor = RGB(255, 255, 255) ' Set white font color
            Else
                Me.Validity.BackColor = RGB(255, 255, 255) ' Set white font color
                Me.Validity.ForeColor = RGB(47, 54, 153)  ' Set tas font color
            End If
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    End If
End Sub
Private Sub RenderRecord_FlightType()
    Dim selectedVendorTable As String
    Dim selectedAirport As Variant
    Dim selectedFlightType As String
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    selectedVendorTable = Me.Vendor
    selectedAirport = Me.AIRPORT.Column(0)
    selectedFlightType = Me.FlightType


    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        ' Find the row that matches the selectedAirport
        'rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        rs.FindFirst "AIRPORT = '" & selectedAirport & "' And FlightType = '" & selectedFlightType & "'"
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.EstimatedTotaltxt.Value = rs("EstimatedTotal").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.Validity.Value = rs("Validity").Value
            Me.Currency.Value = rs("Currency").Value
            ' will be hidden As no need To show them
            Me.FuelType.Value = rs("FuelType").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.Unit.Value = rs("Unit").Value
            Me.InternalNotes.Value = rs("InternalNotes").Value

            rs.Close
            Set rs = Nothing
            Set db = Nothing

            If Me.Validity < date Then
                Me.Validity.BackColor = RGB(255, 0, 0) ' Set red background color
                Me.Validity.ForeColor = RGB(255, 255, 255) ' Set white font color
            Else
                Me.Validity.BackColor = RGB(255, 255, 255) ' Set white font color
                Me.Validity.ForeColor = RGB(47, 54, 153)  ' Set tas font color
            End If
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    End If
End Sub
Private Sub openAddReleaseF()
    DoCmd.OpenForm "FlightSupport_FuelReleaseF_Seperate_New", WindowMode:=acWindowNormal

End Sub
Private Sub FuelStatus_BeforeUpdate(Cancel As Integer)
    Dim Loc_ID As Integer
    Dim rs As DAO.Recordset
    Dim db As DAO.Database


    Loc_ID = Me.AIRPORT.Column(1)


    If (Me.FuelStatus = "CONFIRMED") Or (Me.FuelStatus = "CONFIRMED - STC") Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset("FuelReleaseT_New_V12", dbOpenSnapshot)

        rs.FindFirst "Location_ID = " & Loc_ID

        If rs.NoMatch Then
            MsgBox "No matching Fuel Release found. Please add a release before changing status To confirmed.", vbExclamation + vbOKOnly, "Record Not Found"
            'Me.Undo
            Cancel = True ' Cancel the update
            rs.Close
            Set rs = Nothing
            openAddReleaseF
         Exit Sub

            ' Exit the procedure
        Else
            MsgBox "Status changed successfully."
            Me.FuelStatus.BackColor = RGB(0, 255, 0) ' Green color

        End If

    Else
        Me.FuelStatus.BackColor = RGB(255, 255, 255)
     Exit Sub

    End If


End Sub
Private Sub Vendor_AfterUpdate()
    RenderRecord_Two
End Sub
Private Sub RenderRecord_Two()
    Dim selectedVendorTable As String
    Dim selectedAirport As Variant
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Dim selectedFlightType As String

    selectedVendorTable = Me.Vendor
    selectedAirport = Me.AIRPORT.Column(0)
    selectedFlightType = Me.FlightType


    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        ' Find the row that matches the selectedAirport
        'rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        rs.FindFirst "AIRPORT = '" & selectedAirport & "' And FlightType = '" & selectedFlightType & "'"
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.EstimatedTotaltxt.Value = rs("EstimatedTotal").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.Validity.Value = rs("Validity").Value
            Me.Currency.Value = rs("Currency").Value
            ' will be hidden As no need To show them
            Me.FuelType.Value = rs("FuelType").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.Unit.Value = rs("Unit").Value
            Me.InternalNotes.Value = rs("InternalNotes").Value


            If Me.Validity < date Then
                Me.Validity.BackColor = RGB(255, 0, 0) ' Set red background color
                Me.Validity.ForeColor = RGB(255, 255, 255) ' Set white font color
            Else
                Me.Validity.BackColor = RGB(255, 255, 255) ' Set white font color
                Me.Validity.ForeColor = RGB(47, 54, 153)  ' Set tas font color
            End If
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
End Sub

Private Sub QuoteReset()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    Me.IPA.Value = Null
    Me.Qouted_Vendor.Value = Null
    Me.Vendor.Value = Null
    Me.EstimatedTotaltxt.Value = Null
    Me.FlightType.Value = Null
    Me.Validity.Value = Null
    Me.Currency.Value = Null
    Me.Unit.Value = Null
    Me.InternalNotes.Value = Null
    Me.AIRPORT.Value = Null
    Me.FuelStatus.Value = Null
    RowSourceReset
End Sub
Private Sub RowSourceReset()
    Dim selectedSect_ID As Integer
    Dim LocationsQuery As String

    If IsNull(selectedSect_ID) Then
     Exit Sub
    Else
        selectedSect_ID = Forms("FlightSupport_LASTFORM_New_V12").Controls("FlightSupport_AddUpdateRequestF").Controls("Add_UpdateRequestF").Form.Controls("SectorNocmb").Column(1)
        LocationsQuery = "Select DISTINCT Location, Loc_ID FROM FlightSupport_locationsQ_V12 WHERE Sect_ID = " & selectedSect_ID
        Me.AIRPORT.RowSource = LocationsQuery
        Me.AIRPORT.Requery
    End If

End Sub

Private Sub ResetBtn_Click()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    Me.IPA.Value = Null
    Me.FlightType.Value = Null
    Me.AIRPORT.Value = Null
    'Me.Location_ID.Value = Null
    Me.FuelType.Value = Null
    Me.FuelStatus.Value = Null
    Me.QuoteSentCb.Value = Null
    Me.EstimatedTotaltxt.Value = Null
    ' Not needed
    'RowSourceReset

End Sub

Private Sub Validity_Click()
    InputDateField Validity, "Select a date To use this on your form"
End Sub
Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FuelSupport_FuelQuote", acSaveNo
End Sub
Private Sub AddFuelService_Click()
    'On Error Goto ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If IsNull(Me.AIRPORT) Or IsNull(Me.FuelStatus) Or IsNull(Me.PaymentMethod) Or IsNull(Me.Quantity) Or IsNull(Me.Currency) Or IsNull(Me.vendorName) Or IsNull(Me.AddedRate) Or IsNull(Me.TASPrice) Then
        MsgBox "Add Missing Field.", vbExclamation
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset("FuelBriefT_New", dbOpenDynaset)

        ' Find the existing record based on Location_ID And vendorName
        rs.FindFirst "Location_ID = " & Me.AIRPORT.Column(1) & " And Vendor = '" & Me.vendorName.Value & "'"

        If rs.NoMatch Then ' No existing record found
            rs.AddNew

            ' Assign values To the fields in the table
            rs.Fields("Location_ID").Value = Me.AIRPORT.Column(1)
            rs.Fields("FuelType").Value = Me.FuelType.Value
            rs.Fields("FuelStatus").Value = Me.FuelStatus.Value
            rs.Fields("PaymentMethod").Value = Me.PaymentMethod.Value
            rs.Fields("IPA").Value = Me.IPA.Value
            rs.Fields("EstimatedQuantity").Value = Me.Quantity.Value
            rs.Fields("VendorPrice").Value = Me.BasePrice.Value
            rs.Fields("TASPrice").Value = Me.TASPrice.Value
            rs.Fields("AddedRate").Value = Me.AddedRate.Value
            rs.Fields("Vendor").Value = Me.vendorName.Value
            rs.Fields("FBO").Value = Me.FBO.Value
            If Me.QuoteSentCb.Value = True Then
                rs.Fields("QuoteSent").Value = Me.QuoteSentCb.Value
                rs.Fields("Q_Vendor").Value = Me.Qouted_Vendor.Column(0)
                rs.Fields("FuelQuote_RefNo").Value = Me.Qouted_Vendor.Column(1)
            Else
                rs.Fields("QuoteSent").Value = Me.QuoteSentCb.Value
            End If
            rs.Fields("QuoteSent").Value = Me.QuoteSentCb.Value

            rs.Fields("Taxes").Value = Me.Taxes.Value
            rs.Fields("Notes").Value = Me.Notes.Value
            rs.Fields("Validity").Value = Me.Validity.Value
            rs.Fields("EstimatedTotal").Value = Me.EstimatedTotaltxt.Value
            rs.Fields("Location").Value = Me.AIRPORT.Column(0)
            rs.Fields("Currency").Value = Me.Currency.Value
            rs.Fields("InternalNotes").Value = Me.InternalNotes.Value
            rs.Fields("FlightType").Value = Me.FlightType.Value
            ' Save the New record
            rs.Update

            ' Clean up
            rs.Close
            Set rs = Nothing
            Set db = Nothing

            ' Reset form values
            QuoteReset
            ' Display a message indicating successful addition
            MsgBox "Service added successfully. Please add fuel release once service is confirmed from the vendor."
            Me.Undo
        Else
            ' Clean up
            rs.Close
            Set rs = Nothing
            Set db = Nothing

            MsgBox "Vendor already exists For the same Location_ID.", vbExclamation
        End If

     Exit Sub
    End If

    'ErrorHandler:
    'MsgBox "An unexpected error occurred: " & err.Description, vbCritical, "Error"
    'Me.Undo
    'Exit Sub
End Sub
