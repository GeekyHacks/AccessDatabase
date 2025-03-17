Option Compare Database
Option Explicit
Private Sub ShowCountry()
    Dim selectedValue As Variant
    selectedValue = Me.AIRPORT.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("Airports_DataT", dbOpenSnapshot)

        rs.FindFirst "AirportCode = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.Country.Value = rs("Country").Value
        Else
            Me.Country.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.Country.Value = Null
    End If

 Exit Sub

End Sub

Private Sub AIRPORT_GotFocus()
    MinPrice_MinVendor
    ShowCountry
End Sub

Private Sub Airport_AfterUpdate()
    MinPrice_MinVendor
    ShowCountry
End Sub
Private Sub MinPrice_MinVendor()
    Dim selectedAirport As Variant
    'Dim airportCode As String
    Dim LowestPriceQ As String
    Dim MinVendorPrice As Variant
    Dim AirportIDQ As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    selectedAirport = Me.AIRPORT

    If Not IsNull(selectedAirport) Then
        ' airportCode = selectedAirport
        Set db = CurrentDb
        Set rs = db.OpenRecordset("Select * FROM FuelSupport_NewFuelQ WHERE Airport = '" & selectedAirport & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            Me.AirportID.Value = rs!ID
            Me.LowestVendor.Value = rs!TableName
            Me.Vendor.Value = rs!TableName
            Me.LowestPriceTxt.Value = rs!MinPrice

            RenderRecord
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    End If
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
        'rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.EstimatedTotaltxt.Value = rs("EstimatedTotal").Value
            Me.IPA.Value = rs("IPA").Value
            Me.FuelType.Value = rs("FuelType").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.Validity.Value = rs("Validity").Value
            Me.VendorCmbPrice.Value = rs("VendorPrice").Value
            Me.Currency.Value = rs("Currency").Value
            Me.UNIT.Value = rs("Unit").Value
            Me.InternalNotes.Value = rs("InternalNotes").Value
            Me.ChangeCyclebx.Value = rs("ChangeCycle").Value
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
Private Sub MinPrice_MinVendor_FlightType()
    Dim selectedAirport As Variant
    Dim LowestPriceQ As String
    Dim selectedFlightType As String
    Dim MinVendorPrice As Variant
    Dim AirportIDQ As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    selectedAirport = Me.AIRPORT
    selectedFlightType = Me.FlightType
    If Not IsNull(selectedAirport) Then
        ' airportCode = selectedAirport
        Set db = CurrentDb
        Set rs = db.OpenRecordset("Select * FROM FuelSupport_NewFuelQ WHERE FlightType = '" & selectedFlightType & "' And AIRPORT = '" & selectedAirport & "'", dbOpenSnapshot)
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
Private Sub Command35_Click()
    DoCmd.OpenForm "FuelSupport_updateFuel_New", WindowMode:=acWindowNormal
End Sub

Private Sub CheckBtn_Click()
    DoCmd.OpenForm "FuelSupport_FuelQuote", WindowMode:=acWindowNormal

End Sub

Private Sub Client_AfterUpdate()
    CustomerIDSelection
End Sub
Private Sub FlightType_AfterUpdate()
    MinPrice_MinVendor_FlightType
End Sub
Private Sub RenderRecord_Two()
    Dim selectedVendorTable As String
    Dim selectedAirport As Variant
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Dim selectedFlightType As String

    selectedVendorTable = Me.Vendor
    selectedAirport = Me.AIRPORT
    selectedFlightType = Me.FlightType


    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        ' Find the row that matches the selectedAirport
        'rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        rs.FindFirst "FlightType = '" & selectedFlightType & "' And AIRPORT = '" & selectedAirport & "'"
        If Not rs.NoMatch Then
            Me.ChangeCyclebx.Value = rs("ChangeCycle").Value
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.EstimatedTotaltxt.Value = rs("EstimatedTotal").Value
            Me.IPA.Value = rs("IPA").Value
            Me.FuelType.Value = rs("FuelType").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.Validity.Value = rs("Validity").Value
            Me.VendorCmbPrice.Value = rs("VendorPrice").Value
            Me.Currency.Value = rs("Currency").Value
            Me.UNIT.Value = rs("Unit").Value
            Me.InternalNotes.Value = rs("InternalNotes").Value
            Me.ChangeCyclebx.Value = rs("ChangeCycle").Value
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
    selectedAirport = Me.AIRPORT
    selectedFlightType = Me.FlightType


    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        ' Find the row that matches the selectedAirport
        'rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        rs.FindFirst "FlightType = '" & selectedFlightType & "' And AIRPORT = '" & selectedAirport & "'"
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.ChangeCyclebx.Value = rs("ChangeCycle").Value
            Me.BasePrice.Value = rs("VendorPrice").Value
            Me.Notes.Value = rs("Notes").Value
            Me.Taxes.Value = rs("Taxes").Value
            Me.EstimatedTotaltxt.Value = rs("EstimatedTotal").Value
            Me.IPA.Value = rs("IPA").Value
            Me.FuelType.Value = rs("FuelType").Value
            Me.FlightType.Value = rs("FlightType").Value
            Me.Validity.Value = rs("Validity").Value
            Me.VendorCmbPrice.Value = rs("VendorPrice").Value
            Me.Currency.Value = rs("Currency").Value
            Me.UNIT.Value = rs("Unit").Value
            Me.InternalNotes.Value = rs("InternalNotes").Value
            Me.ChangeCyclebx.Value = rs("ChangeCycle").Value

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
Private Sub CustomerIDSelection()
    On Error Goto ErrorHandler

        Dim selectedValue As Variant
        selectedValue = Me.Client.Column(1) ' Get the selected value from the combo box

        If Not IsNull(selectedValue) Then
            Dim rs As DAO.Recordset
            Set rs = CurrentDb.OpenRecordset("CustomersT", dbOpenSnapshot)

            ' Assuming ID is a numeric field, remove the single quotes around selectedValue
            rs.FindFirst "ID = " & selectedValue
            If Not rs.NoMatch Then
                Me.ClientIDtxt.Value = rs("ID").Value
            Else
                Me.ClientIDtxt.Value = Null
            End If

            rs.Close
            Set rs = Nothing
        Else
            Me.ClientIDtxt.Value = Null
        End If

     Exit Sub

 ErrorHandler:
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        If Not rs Is Nothing Then
            rs.Close
            Set rs = Nothing
        End If
End Sub

Private Sub Command36_Click()
    DoCmd.OpenForm "FuelSupport_UpdateFuel_New", WindowMode:=acWindowNormal
End Sub


Private Sub Command659_Click()
    Me.Refresh
End Sub

Private Sub cmdGenerate_Click()
    Dim RefValue As Integer
    Dim rst As DAO.Recordset
    RefValue = Nz(DLookup("Quote_ID", "TRefNo"), 0)

    Me.Quote_RefNotxt = "TASFQ" & Format(RefValue + 7, "0000") & Format(date, "yy")
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
    If IsNull(Me.Quote_RefNotxt) Or Me.Quote_RefNotxt = "" Then
        cmdGenerate_Click
    End If
    ' Set focus To the first field For data entry
    'Me.AIRPORT.SetFocus
End Sub
Private Sub Vendor_AfterUpdate()
    RenderRecord_Two
End Sub
Private Sub QuoteReset()
    Me.BasePrice.Value = Null
    Me.Notes.Value = Null
    Me.Taxes.Value = Null
    Me.IPA.Value = Null
    Me.AIRPORT.Value = Null
    Me.ChangeCyclebx.Value = Null

    'RowSourceReset
End Sub
Private Sub RowSourceReset()
    Dim AirportQuery As String
    AirportQuery = "Select Airport FROM FuelSupport_PriceQ ORDER BY Airport ASC"
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
    Me.Quote_RefNotxt.Value = Null
    Me.ChangeCyclebx.Value = Null
    ' RowSourceReset

End Sub

Private Sub Validity_Click()
    InputDateField Validity, "Select a date To use this on your form"
End Sub
Private Sub CloseBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FuelSupport_FuelQuote", acSaveNo
End Sub
Private Sub AddFuelQuote_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim FatherForm As Form

    Set FatherForm = Forms("FuelSupport_QuoteRequestF_V1").Form

    If IsNull(Me.BasePrice.Value) Or IsNull(Me.Validity.Value) Or IsNull(Me.Client.Value) Or IsNull(Me.AIRPORT.Value) _
        Or IsNull(Me.TASPrice.Value) Or IsNull(Me.AddedRate.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    End If

    If IsNull(Me.Quote_RefNotxt.Value) Then
        newQuoteBtn_Click
    End If

    If CheckIFQuoteIsAlreadyAdded() Then
        MsgBox "Quote For the same airport already exists."
     Exit Sub
    End If


    Set db = CurrentDb
    Set rs = db.OpenRecordset("FuelSupport_FuelQuotationsT", dbOpenDynaset)

    rs.AddNew
    'MsgBox Me.BasePrice.Value
    rs.Fields("VendorPrice").Value = Me.BasePrice.Value
    ' MsgBox Me.Notes.Value
    rs.Fields("Notes").Value = Me.Notes.Value
    ' rs.Fields("LowCurrency").Value = Me.LowCurrency.Value
    rs.Fields("Taxes").Value = Me.Taxes.Value
    rs.Fields("EstimatedTotal").Value = Me.EstimatedTotaltxt.Value
    rs.Fields("TASEstimatedTotal").Value = Me.TASEstimatedTotal.Value
    rs.Fields("IPA").Value = Me.IPA.Value
    rs.Fields("FuelType").Value = Me.FuelType.Value
    rs.Fields("FlightType").Value = Me.FlightType.Value
    rs.Fields("Airport").Value = Me.AIRPORT.Value
    rs.Fields("Currency").Value = Me.Currency.Value
    rs.Fields("Unit").Value = Me.UNIT.Value
    rs.Fields("Country").Value = Me.Country.Value
    rs.Fields("EffDate").Value = Me.EffDate.Value
    rs.Fields("Customer").Value = Me.Client.Value
    rs.Fields("Q_TASPrice").Value = Me.TASPrice.Value
    rs.Fields("Quote_RefNo").Value = Me.Quote_RefNotxt.Value
    rs.Fields("AddedRate").Value = Me.AddedRate.Value
    rs.Fields("Vendor").Value = Me.vendorName.Value ' Add this line
    rs.Fields("Validity").Value = Me.Validity.Value
    rs.Fields("ChangeCycle").Value = Me.ChangeCyclebx.Value
    rs.Fields("QuoteRequestRefNo").Value = Me.R_RefNo.Value
    rs.Fields("MinUpliftFees").Value = Me.UpliftFees.Value
    rs.Fields("MinUplift").Value = Me.MinUplift.Value
    rs.Fields("User").Value = Me.Usertxt.Value
    ' Save the New record
    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset the form fields
    QuoteReset
    'FatherForm.Controls("FuelSupport_LocationsListF").Form.Controls("QuoteRefNotxt").Requery

    MsgBox "Data added successfully.", vbInformation, "Success"
    Me.Undo
End Sub
Private Function CheckIFQuoteIsAlreadyAdded() As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim currentQuoteRefNo As Variant
    Dim selectedAirport As Variant

    currentQuoteRefNo = Me.Quote_RefNotxt.Value
    selectedAirport = Me.AIRPORT.Value

    Set db = CurrentDb
    Set rs = db.OpenRecordset("FuelSupport_FuelQuotationsT", dbOpenDynaset)

    If Not rs.EOF Then
        rs.MoveFirst ' Move To the first record in the recordset

        ' Loop through the recordset
        Do Until rs.EOF
            If rs("Quote_RefNo").Value = currentQuoteRefNo And rs("Airport").Value = selectedAirport And rs("FlightType").Value = Me.FlightType.Value Then
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
Private Sub ChangeQuoteStatus()
    On Error Goto ErrorHandler
        Dim FatherForm As Form
        Dim db As DAO.Database
        Dim rsX As DAO.Recordset
        Dim rsY As DAO.Recordset
        Dim strSQL As String
        Dim QRequestRefNo As String ' Assuming QRequestRefNo is a string; change To Long If it's a number

        ' Initialize database And recordsets
        Set db = CurrentDb
        Set FatherForm = Forms("FuelSupport_QuoteRequestF_V1").Form

        ' Get the QuoteRequestRefNo value (replace this With your actual logic To Get the value)
        QRequestRefNo = Me.R_RefNo ' Example: Get the value from a form control

        ' Open Table Y (FuelSupport_FuelQuotationsT) With a filter For the specific QuoteRequestRefNo
        Set rsY = db.OpenRecordset("Select * FROM FuelSupport_FuelQuotationsT WHERE QuoteRequestRefNo = '" & QRequestRefNo & "'", dbOpenDynaset)

        ' Check If Table Y has records
        If rsY.EOF And rsY.BOF Then
            MsgBox "No records found in Table Y For QuoteRequestRefNo: " & QRequestRefNo, vbExclamation
         Exit Sub
        End If

        ' Loop through Table Y And update Table X For matching locations
        rsY.MoveFirst
        Do While Not rsY.EOF
            ' Open Table X With a filter For the current location And QuoteRequestRefNo
            strSQL = "Select * FROM FuelSupport_RequestedLocationsT_V1 WHERE QuoteRequestRefNo = '" & QRequestRefNo & "' And Location = '" & rsY!AIRPORT & "'"
            Set rsX = db.OpenRecordset(strSQL, dbOpenDynaset)

            ' If a matching record is found, update Table X
            If Not rsX.EOF And Not rsX.BOF Then
                rsX.MoveFirst
                Do While Not rsX.EOF
                    rsX.Edit
                    rsX!LocationQuoteStatus = "Sent"
                    rsX!QuoteRefNo = rsY!Quote_RefNo '' Update values in Table X With QuoteStatus from Table Y
                    rsX.Update
                    rsX.MoveNext
                Loop
            End If

            ' Close the Table X recordset For the current location
            rsX.Close
            Set rsX = Nothing

            ' Move To the Next record in Table Y
            rsY.MoveNext
        Loop

        ' Clean up
        rsY.Close
        Set rsY = Nothing
        Set db = Nothing

        MsgBox "Record has been updated successfully For QuoteRequestRefNo: " & QRequestRefNo, vbInformation
     Exit Sub

        FatherForm.Controls("FuelSupport_LocationsListF").Form.Requery

 ErrorHandler:
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        ' Clean up in Case of error
        If Not rsX Is Nothing Then rsX.Close
            If Not rsY Is Nothing Then rsY.Close
                Set rsX = Nothing
                Set rsY = Nothing
                Set db = Nothing
End Sub

Private Sub PrintQuote_Click()
    If Not IsNull(Me.Quote_RefNotxt) Then
        Dim strWhere As String
        Dim InLeng As Long

        Const conJetDate = "\#dd\/mm\/yyyy\#"

        If Not IsNull(Me.Quote_RefNotxt) Then
            strWhere = strWhere & "([Quote_RefNo] Like '*" & Me.Quote_RefNotxt & "*') And "
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
Private Sub EmailQuoteBtn_Click()
    On Error Goto ErrorHandler

        ' Validate input
        If IsNull(Me.Quote_RefNotxt) Or Me.Quote_RefNotxt = "" Then
            MsgBox "Quote Reference Number is required.", vbExclamation
         Exit Sub
        End If

        If MsgBox("Press Yes To view the Quote", vbYesNo + vbQuestion) = vbYes Then
            PrintQuote_Click ' Opens report modally; code pauses here Until report is closed
        Else
         Exit Sub
        End If


        ' Sanitize input
        Dim sanitizedQuoteRefNo As String
        sanitizedQuoteRefNo = Replace(Me.Quote_RefNotxt, "'", "''")

        ' Construct filter
        Dim strWhere As String
        strWhere = "([Quote_RefNo] Like '*" & sanitizedQuoteRefNo & "*') And "
        strWhere = Left$(strWhere, Len(strWhere) - 5)

        ' Apply filter To the form
        Me.Filter = strWhere
        Me.FilterOn = True

        ' Extract values from the report
        Dim AIRPORT As String
        Dim clientID As String
        Dim recipientEmail As String
        Dim SecondaryEmail As String
        Dim fileName As String
        Dim rs As DAO.Recordset

        ' Use parameterized query To prevent SQL injection
        Dim qdf As DAO.QueryDef
        Set qdf = CurrentDb.CreateQueryDef("", "Select DISTINCT Airport FROM FuelSupport_FuelQuotationsT WHERE [Quote_RefNo] Like @QuoteRefNo")
        qdf.Parameters("@QuoteRefNo").Value = "*" & Me.Quote_RefNotxt & "*"
        Set rs = qdf.OpenRecordset(dbOpenSnapshot)

        AIRPORT = ""
        Do While Not rs.EOF
            AIRPORT = AIRPORT & Left(rs!AIRPORT, 4) & "_"
            rs.MoveNext
        Loop

        ' Remove trailing underscores
        If Len(AIRPORT) > 0 Then AIRPORT = Left(AIRPORT, Len(AIRPORT) - 1)

            rs.Close
            Set rs = Nothing

            ' Get client email
            clientID = Me.ClientIDtxt.Value
            recipientEmail = Nz(DLookup("PrimaryEmail", "CustomersT", "ID = " & clientID), "")
            SecondaryEmail = Nz(DLookup("SecondaryEmail", "CustomersT", "ID = " & clientID), "")
            LogToFile "Client email: " & recipientEmail
            LogToFile "Client secondary email: " & SecondaryEmail

            ' Construct And sanitize file name
            fileName = SanitizeFileName(Me.Quote_RefNotxt & " " & AIRPORT & " " & clientID & ".pdf")
            LogToFile "File name: " & fileName

            ' Open file dialog To Select path
            Dim fd As fileDialog
            Set fd = Application.fileDialog(msoFileDialogFolderPicker)
            Dim filePath As String

            With fd
                .Title = "Select Folder"
                .AllowMultiSelect = False
                If .Show = -1 Then
                    filePath = .SelectedItems(1) & "\" & fileName
                    LogToFile "Selected file path: " & filePath
                Else
                    MsgBox "No folder selected. Operation cancelled.", vbExclamation
                 Exit Sub
                End If
            End With

            ' Export report To PDF
            DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
            DoCmd.OutputTo acOutputReport, "FuelSupport_FuelQuote", acFormatPDF, filePath, False
            DoCmd.Close acReport, "FuelSupport_FuelQuote"
            LogToFile "Report exported To: " & filePath

            ' Send email If recipient is found
            If recipientEmail <> "" Then
                Dim olApp As Object
                Dim olMailItem As Object
                Dim currentUserEmail As String
                Dim signature As String

                ' Initialize Outlook application
                Set olApp = CreateObject("Outlook.Application")
                Set olMailItem = olApp.CreateItem(0) ' 0 = olMailItem

                ' Get the current user's email address
                currentUserEmail = olApp.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                LogToFile "Current user email: " & currentUserEmail

                ' Get the current user's signature
                signature = GetOutlookSignature(currentUserEmail, olMailItem)
                'LogToFile "Signature retrieved: " & signature

                ' Set email properties
                With olMailItem
                    .SentOnBehalfOfName = "occ@tahseenaviation.com" ' Send on behalf of OCC
                    .To = recipientEmail
                    If SecondaryEmail <> "" Then
                        .To = .To & ";" & SecondaryEmail ' Add secondary email To the "To" field
                    End If
                    .CC = currentUserEmail & ";occ@tahseenaviation.com" ' CC current user And OCC
                    .Subject = "Tahseen Fuel Quote: #" & Me.R_RefNo & " For " & AIRPORT
                    .Attachments.aDD filePath

                    ' Build the email body With the signature
                    .HTMLBody = "Dear On Duty, " & "<br><br>" & _
                    "Thank you For your inquiry And For choosing Tahseen Aviation Services." & "<br><br>" & _
                    "Tahseen Fuel Quote #" & Me.Quote_RefNotxt & " For " & AIRPORT & " has been forwarded To your email address. Please confirm your acceptance To attached quote To proceed With your order." & "<br><br>" & _
                    "Please Do Not hesitate To contact us If you have any further fuel requests Or inquiries." & "<br><br>" & _
                    "<br>" & signature ' Append the signature

                    ' Send the email after confirmation
                    If MsgBox("Are you sure you want To send this email?", vbYesNo + vbQuestion) = vbYes Then
                        .Send
                        MsgBox "Email sent successfully With the attached quote.", vbInformation, "Email Sent"
                        LogToFile "Email sent successfully."
                        ChangeQuoteStatus
                    Else
                        MsgBox "Email Not sent.", vbInformation, "Operation Cancelled"
                        LogToFile "Email sending cancelled by user."
                    End If
                End With
            Else
                MsgBox "Client email Not found.", vbExclamation, "Email Error"
                LogToFile "Client email Not found."
            End If

         Exit Sub

 ErrorHandler:
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
            LogToFile "Error in EmailQuoteBtn_Click: " & Err.Description
End Sub
Function GetOutlookSignature(currentUserEmail As String, Byref olMailItem As Object) As String
    On Error Goto ErrorHandler
        ' Declare the olByValue constant
        Const olByValue = 1

        Dim signature As String
        Dim signaturePath As String
        Dim fso As Object
        Dim ts As Object
        Dim imgFolderPath As String
        Dim imgFile As Object
        Dim imgFiles As Object
        Dim imgPattern As String
        Dim imgSrc As String
        Dim cid As String
        Dim regex As Object, matches As Object

        ' Get the path To the Outlook signature folder
        signaturePath = Environ("AppData") & "\Microsoft\Signatures\"
        LogToFile "Signature folder path: " & signaturePath

        ' Construct the signature file name dynamically
        Dim signatureName As String
        signatureName = "OCC (" & currentUserEmail & ")" ' Format: OCC (currentuseremail)
        LogToFile "Using signature file name: " & signatureName

        ' Read the signature file
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(signaturePath & signatureName & ".htm") Then
            Set ts = fso.OpenTextFile(signaturePath & signatureName & ".htm", 1) ' 1 = ForReading
            signature = ts.ReadAll
            ts.Close
        Else
            GetOutlookSignature = ""
            LogToFile "Signature file Not found: " & signaturePath & signatureName & ".htm"
         Exit Function
        End If

        ' Get the path To the signature's image folder
        imgFolderPath = signaturePath & signatureName & "_files\"
        LogToFile "Image folder path: " & imgFolderPath

        ' Replace image references With CID attachments
        If fso.FolderExists(imgFolderPath) Then
            Set imgFiles = fso.GetFolder(imgFolderPath).Files
            Set regex = CreateObject("VBScript.RegExp")
            regex.Global = True
            regex.IgnoreCase = True

            For Each imgFile In imgFiles
                'imgPattern = "<img[^>]*src=""" & imgFile.Name & """[^>]*>"
                imgPattern = "<img[^>]*src=""([^""]*" & Replace(imgFile.Name, ".", "\.") & """)""[^>]*>" ' Modified pattern

                regex.Pattern = imgPattern
                If regex.test(signature) Then
                    Set matches = regex.Execute(signature)
                    If matches.Count > 0 Then
                        ' Attach the image To the email And generate a CID
                        cid = "img" & imgFile.Name
                        olMailItem.Attachments.aDD imgFile.Path, 1, 0, cid

                        ' Replace the image source With the CID reference
                        imgSrc = "cid:" & cid
                        signature = Replace(signature, matches(0).Value, "<img src=""" & imgSrc & """>") ' Replace the entire matched string
                        ' LogToFile "Replaced image source: " & imgFile.Name & " With CID: " & imgSrc
                    End If
                End If
            Next imgFile
            Set regex = Nothing
        End If

        ' Clean up
        Set ts = Nothing
        Set fso = Nothing

        GetOutlookSignature = signature
     Exit Function

 ErrorHandler:
        ' Fallback: Return an empty string If the signature cannot be retrieved
        GetOutlookSignature = ""
        LogToFile "Error retrieving Outlook signature: " & Err.Description
End Function
Function BinaryToBase64(binaryData) As String
    Dim xmlDoc As Object
    Dim xmlNode As Object

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")

    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = binaryData
    BinaryToBase64 = xmlNode.Text

    Set xmlNode = Nothing
    Set xmlDoc = Nothing
End Function

' Function To sanitize file names
Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer

    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i

    SanitizeFileName = fileName
End Function

' Function To log messages securely
Private Sub LogToFile(Message As String)
    Dim filePath As String
    Dim fileNumber As Integer

    filePath = "E:\ent\DebugLog.txt" ' Change To a secure location
    fileNumber = FreeFile

    Open filePath For Append As #fileNumber
    Print #fileNumber, Now() & " - " & Message
    Close #fileNumber
End Sub

