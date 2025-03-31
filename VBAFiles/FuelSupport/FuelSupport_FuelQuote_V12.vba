Option Compare Database
Private Sub ShowCountry()
    Dim selectedValue As Variant
    selectedValue = Me.AIRPORT.value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("Airports_DataT", dbOpenSnapshot)

        rs.FindFirst "AirportCode = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.Country.value = rs("Country").value
        Else
            Me.Country.value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.Country.value = Null
    End If

 Exit Sub

End Sub
Private Sub AddNewVendorTablebtn_Click()
    On Error GoTo ErrorHandler

    ' Define the source table name
    Dim strSourceTable As String
    strSourceTable = "NewFuelT" ' Replace with your source table name

    ' Get the new table name from a text field
    Dim strNewTable As String
    strNewTable = Me.NewVendorTable.value ' Replace with your text field name

    ' Validate the new table name
    If IsNull(strNewTable) Or strNewTable = "" Then
        MsgBox "New table name cannot be empty!", vbExclamation
        Exit Sub
    End If

    ' Sanitize and validate the new table name
    strNewTable = Replace(strNewTable, "'", "''")
    strNewTable = Replace(strNewTable, ";", "")
    strNewTable = Replace(strNewTable, "--", "")
    strNewTable = Replace(strNewTable, "/*", "")
    strNewTable = Replace(strNewTable, "*/", "")

    If Not IsValidTableName(strNewTable) Then
        MsgBox "Invalid table name! Only alphanumeric characters and underscores are allowed.", vbExclamation
        Exit Sub
    End If

    If Len(strNewTable) > 64 Then
        MsgBox "Table name cannot exceed 64 characters!", vbExclamation
        Exit Sub
    End If

    ' Append "_Lo" to the local table name
    Dim strLocalTable As String
    strLocalTable = strNewTable & "_Lo"

    ' Check if the local table already exists and delete it if it does
    Dim db As DAO.Database
    Set db = CurrentDb()

    If TableExists(strLocalTable) Then
        On Error Resume Next
        db.TableDefs.Delete strLocalTable
        On Error GoTo ErrorHandler
    End If

    ' Use TransferDatabase to import the table structure as a local table
    DoCmd.TransferDatabase acImport, "Microsoft Access", CurrentDb.Name, acTable, strSourceTable, strLocalTable

    ' Ensure the new table is empty (delete all data)
    db.Execute "DELETE FROM " & strLocalTable, dbFailOnError

    ' Add the new table name to the FuelTablesList column in the AviationServicesT table
    db.Execute "INSERT INTO AviationServicesT (FuelTablesList) VALUES ('" & strNewTable & "')", dbFailOnError

    ' Refresh the table links to ensure proper synchronization
    RefreshTableLinks

    MsgBox "Table structure copied, renamed, and added to FuelTablesList successfully!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    LogError Err.Number, Err.Description
    Exit Sub
End Sub

' Helper function to check if a table exists
Private Function TableExists(TableName As String) As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Set db = CurrentDb()

    TableExists = False
    For Each tdf In db.TableDefs
        If tdf.Name = TableName Then
            TableExists = True
            Exit For
        End If
    Next tdf
End Function

' Helper function to validate table names
Private Function IsValidTableName(TableName As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(TableName)
        Dim char As String
        char = Mid(TableName, i, 1)
        If Not (char Like "[A-Za-z0-9_]") Then
            IsValidTableName = False
            Exit Function
        End If
    Next i
    IsValidTableName = True
End Function

' Helper function to log errors
Private Sub LogError(errNumber As Long, errDescription As String)
    Dim db As DAO.Database
    Set db = CurrentDb()
    db.Execute "INSERT INTO ErrorLog (ErrorNumber, ErrorDescription, UserName, Timestamp) VALUES (" & _
               errNumber & ", '" & Replace(errDescription, "'", "''") & "', '" & Environ("Username") & "', #" & Now & "#)", dbFailOnError
End Sub

' Helper function to refresh table links
Private Sub RefreshTableLinks()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Set db = CurrentDb()

    For Each tdf In db.TableDefs
        If tdf.Connect <> "" Then
            tdf.RefreshLink
        End If
    Next tdf
End Sub

Private Sub Airport_AfterUpdate()
    MinPrice_MinVendor
    ShowCountry
End Sub

Private Sub MinPrice_MinVendor()
    Dim selectedAirport As Variant
    Dim MinVendorPrice As Variant
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim criteria As String
    
    On Error GoTo ErrorHandler
    
    selectedAirport = Me.AIRPORT
    
    If Not IsNull(selectedAirport) Then
        ' Get the minimum price for the selected airport
        MinVendorPrice = DMin("MinPrice", "FuelSupport_ValidFuelPricesT_V12", "Airport = '" & selectedAirport & "'")
        
        If Not IsNull(MinVendorPrice) Then
            Set db = CurrentDb
            criteria = "SELECT Airport, ID, TableName, MinPrice " & _
                       "FROM FuelSupport_ValidFuelPricesT_V12 " & _
                       "WHERE Airport = '" & selectedAirport & "' " & _
                       "AND MinPrice = " & MinVendorPrice & ";"
            
            Set rs = db.OpenRecordset(criteria, dbOpenSnapshot)
            
            If Not rs.EOF Then
                Me.AirportID.value = rs!ID
                Me.LowestVendor.value = Left(rs!TableName, Len(rs!TableName) - 5)
                Me.Vendor.value = rs!TableName
                Me.LowestPriceTxt.value = rs!MinPrice
                
                RenderRecord ' Ensure this subroutine is error-free
            Else
                MsgBox "No vendor found with the lowest price.", vbExclamation, "No Record"
            End If
        Else
            MsgBox "No prices available for the selected airport.", vbExclamation, "No Data"
        End If
    End If
       
        If Not Me.ValidFuelPrices.Form Is Nothing Then
            Me.ValidFuelPrices.Form.RecordSource = "FuelSupport_ValidFuelPricesT_V12"
            Me.ValidFuelPrices.Form.Filter = "Airport = '" & selectedAirport & "'"
            Me.ValidFuelPrices.Form.FilterOn = True
        End If
        
Cleanup:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume Cleanup
End Sub

Private Sub RenderRecord()
    Dim selectedVendorTable As String
    Dim selectedAirport As Variant
    Dim rs As DAO.Recordset
    Dim db As DAO.Database


    selectedVendorTable = Me.Vendor
    selectedAirport = Me.AirportID.value



    If IsNull(selectedAirport) Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        ' Find the row that matches the selectedAirport
        'rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        rs.FindFirst "ID = " & selectedAirport & ""
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.BasePrice.value = rs("VendorPrice").value
            Me.Notes.value = rs("Notes").value
            Me.Taxes.value = rs("Taxes").value
            Me.EstimatedTotaltxt.value = rs("EstimatedTotal").value
            Me.IPA.value = rs("IPA").value
            Me.FuelType.value = rs("FuelType").value
            Me.FlightType.value = rs("FlightType").value
            Me.Validity.value = rs("Validity").value
            Me.VendorCmbPrice.value = rs("VendorPrice").value
            Me.Currency.value = rs("Currency").value
            Me.UNIT.value = rs("Unit").value
            Me.InternalNotes.value = rs("InternalNotes").value
            Me.ChangeCyclebx.value = rs("ChangeCycle").value
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
        Set rs = db.OpenRecordset("Select * FROM FuelSupport_ValidFuelPricesT_V12 WHERE FlightType = '" & selectedFlightType & "' And AIRPORT = '" & selectedAirport & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            Me.AirportID.value = rs!ID
            Me.LowestVendor.value = Left(rs!TableName, Len(rs!TableName) - 5)
            Me.Vendor.value = rs!TableName
            Me.LowestPriceTxt.value = rs!MinPrice


            RenderRecord_FlightType
        Else
            MsgBox "No matching records found.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    End If
End Sub

Private Sub Command35_Click()
    DoCmd.OpenForm "FuelSupport_AddFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command36_Click()
    DoCmd.OpenForm "FuelSupport_updateFuel_New", WindowMode:=acWindowNormal
End Sub

Private Sub EffDate_Click()
Me.EffDate.value = Null
InputDateField EffDate, "Select a date To use this on your form"
 
End Sub

Private Sub FlightType_AfterUpdate()
    MinPrice_MinVendor_FlightType
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
            Me.Filter = strWhe
            Me.FilterOn = True
            DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
        End If
    Else
        DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
    End If
End Sub

Private Sub Command659_Click()
    Me.Refresh
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
            Me.ChangeCyclebx.value = rs("ChangeCycle").value
            Me.BasePrice.value = rs("VendorPrice").value
            Me.Notes.value = rs("Notes").value
            Me.Taxes.value = rs("Taxes").value
            Me.EstimatedTotaltxt.value = rs("EstimatedTotal").value
            Me.IPA.value = rs("IPA").value
            Me.FuelType.value = rs("FuelType").value
            Me.FlightType.value = rs("FlightType").value
            Me.Validity.value = rs("Validity").value
            Me.VendorCmbPrice.value = rs("VendorPrice").value
            Me.Currency.value = rs("Currency").value
            Me.UNIT.value = rs("Unit").value
            Me.InternalNotes.value = rs("InternalNotes").value
            Me.ChangeCyclebx.value = rs("ChangeCycle").value
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
    selectedAirport = Me.AirportID.value
    selectedFlightType = Me.FlightType


    If IsNull(selectedAirport) Or selectedAirport = "" Then
        MsgBox "Please Select an airport.", vbExclamation + vbOKOnly, "Airport Selection"
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset(selectedVendorTable, dbOpenSnapshot)
        ' Find the row that matches the selectedAirport
        'rs.FindFirst "AIRPORT = '" & selectedAirport & "'"
        'rs.FindFirst "FlightType = '" & selectedFlightType & "' And AIRPORT = '" & selectedAirport & "'"
         rs.FindFirst "ID = " & selectedAirport & ""
        If Not rs.NoMatch Then
            ' Assign values To the fields in the form`
            Me.ChangeCyclebx.value = rs("ChangeCycle").value
            Me.BasePrice.value = rs("VendorPrice").value
            Me.Notes.value = rs("Notes").value
            Me.Taxes.value = rs("Taxes").value
            Me.EstimatedTotaltxt.value = rs("EstimatedTotal").value
            Me.IPA.value = rs("IPA").value
            Me.FuelType.value = rs("FuelType").value
            Me.FlightType.value = rs("FlightType").value
            Me.Validity.value = rs("Validity").value
            Me.VendorCmbPrice.value = rs("VendorPrice").value
            Me.Currency.value = rs("Currency").value
            Me.UNIT.value = rs("Unit").value
            Me.InternalNotes.value = rs("InternalNotes").value
            Me.ChangeCyclebx.value = rs("ChangeCycle").value

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
    Me.AIRPORT.SetFocus
End Sub

Private Sub QoutedActualListbtn_Click()
    DoCmd.OpenForm "FuelSupport_QuoteActualPrice_V12", WindowMode:=acWindowNormal
End Sub
Private Sub UpdateFuelPricesQBtn_Click()
    On Error GoTo ErrorHandler
    
    ' Show processing message
    DoCmd.Hourglass True
    Me.Refresh
    
    ' Process the fuel data
    If CreateDynamicFuelQuery() Then
        Call CreateValidFuelPricesTable
        MsgBox "Fuel prices updated successfully!", vbInformation
    Else
        MsgBox "Failed to update fuel prices. Please check the log file.", vbExclamation
    End If
    
Cleanup:
    DoCmd.Hourglass False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in UpdateFuelPricesQBtn_Click: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
Private Sub Vendor_AfterUpdate()
    RenderRecord_Two
End Sub
Private Sub AddUser()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.Usertxt = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
Private Sub Form_Load()
    Dim AddUser As Boolean
    Dim FinanceUser As Boolean
    Dim Trainee As Boolean
    Dim UserID As Integer
    Dim fw As New clFormWindow

    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.Usertxt = DLookup("UserName", "UsersT", "ID =" & UserID & "")
    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
   
    RowSourceReset

    UserID = Forms!loggedInUser!txtUserID
    FinanceUser = DLookup("FinanceDep", "UsersT", "ID =" & UserID)
    Trainee = DLookup("Trainee", "UsersT", "ID =" & UserID)
    AddUser = DLookup("CreateUser", "UsersT", "ID =" & UserID)
    
    If Not AddUser Then
            Me.NewVendorTable.Enabled = False
            Me.AddNewVendorTablebtn.Enabled = False
            Me.UpdateFuelPricesQBtn.Enabled = False
        Else
            Me.UpdateFuelPricesQBtn.Enabled = True
            Me.NewVendorTable.Enabled = True
            Me.AddNewVendorTablebtn.Enabled = True
            
    End If

    If FinanceUser Then
            Me.NewVendorTable.Enabled = False
            Me.AddNewVendorTablebtn.Enabled = False
            Me.UpdateFuelPricesQBtn.Enabled = False
    End If

    If Trainee Then
            Me.NewVendorTable.Enabled = False
            Me.AddNewVendorTablebtn.Enabled = False
            Me.UpdateFuelPricesQBtn.Enabled = False
    End If

End Sub
Private Sub QuoteReset()
    Me.BasePrice.value = Null
    Me.Notes.value = Null
    Me.Taxes.value = Null
    Me.IPA.value = Null
    Me.AIRPORT.value = Null
    Me.ChangeCyclebx.value = Null

    RowSourceReset
End Sub
Private Sub RowSourceReset()
    Dim AirportQuery As String
    Dim ValidPricesForm As Form

    AirportQuery = "SELECT ID, Airport FROM FuelSupport_ValidFuelPricesT_V12 ORDER BY Airport ASC"
    ' Set the RowSource Property of the AIRPORT combo box To the dynamic query
    Me.AIRPORT.RowSource = AirportQuery

    ' Requery the AIRPORT combo box To reflect the updated options
    Me.AIRPORT.Requery
    
    ' First create the base query
    If CreateDynamicFuelQuery() Then
     '  Set ValidPricesForm = Me.ValidFuelPrices.Form
       'ValidPricesForm.RecordSource = ""
    End If
    
    
End Sub
Private Sub ResetBtn_Click()
    Me.BasePrice.value = Null
    Me.Notes.value = Null
    Me.Taxes.value = Null
    Me.IPA.value = Null
    Me.FlightType.value = Null
    Me.AIRPORT.value = Null
    Me.Quote_RefNotxt.value = Null
    Me.ChangeCyclebx.value = Null
    RowSourceReset

End Sub

Private Sub Validity_Click()
   ' InputDateField Validity, "Select a date To use this on your form"
End Sub
Private Sub CloseBtn_Click()
    Me.Undo
    Me.ValidFuelPrices.Form.RecordSource = ""
    DoCmd.Close acForm, "FuelSupport_FuelQuote_V12", acSaveNo
End Sub
Private Sub AddFuelQuote_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If IsNull(Me.BasePrice.value) Or IsNull(Me.Validity.value) Or IsNull(Me.Client.value) Or IsNull(Me.AIRPORT.value) _
        Or IsNull(Me.TASPrice.value) Or IsNull(Me.AddedRate.value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    End If

    If IsNull(Me.Quote_RefNotxt.value) Then
        newQuoteBtn_Click
    End If

    If CheckIFQuoteIsAlreadyAdded() Then
        MsgBox "Quote For the same airport already exists."
     Exit Sub
    End If


    Set db = CurrentDb
    Set rs = db.OpenRecordset("FuelSupport_FuelQuotationsT", dbOpenDynaset)

    rs.AddNew
    rs.Fields("VendorPrice").value = Me.BasePrice.value
    rs.Fields("Notes").value = Me.Notes.value
    rs.Fields("LowCurrency").value = Me.LowCurrency.value
    rs.Fields("Taxes").value = Me.Taxes.value
    rs.Fields("EstimatedTotal").value = Me.EstimatedTotaltxt.value
    rs.Fields("TASEstimatedTotal").value = Me.TASEstimatedTotal.value
    rs.Fields("IPA").value = Me.IPA.value
    rs.Fields("FuelType").value = Me.FuelType.value
    rs.Fields("FlightType").value = Me.FlightType.value
    rs.Fields("Airport").value = Me.AIRPORT.value
    rs.Fields("Currency").value = Me.Currency.value
    rs.Fields("Unit").value = Me.UNIT.value
    rs.Fields("Country").value = Me.Country.value
    rs.Fields("EffDate").value = Me.EffDate.value
    rs.Fields("Customer").value = Me.Client.value
    rs.Fields("Q_TASPrice").value = Me.TASPrice.value
    rs.Fields("Quote_RefNo").value = Me.Quote_RefNotxt.value
    rs.Fields("AddedRate").value = Me.AddedRate.value
    rs.Fields("Vendor").value = Me.vendorName.value ' Add this line
    rs.Fields("Validity").value = Me.Validity.value
    rs.Fields("ChangeCycle").value = Me.ChangeCyclebx.value
    rs.Fields("FlightRequest_RefNo").value = Me.R_RefNo.value
    rs.Fields("MinUpliftFees").value = Me.UpliftFees.value
    rs.Fields("MinUplift").value = Me.MinUplift.value
    rs.Fields("User").value = Me.Usertxt.value
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

    currentQuoteRefNo = Me.Quote_RefNotxt.value
    selectedAirport = Me.AIRPORT.value

    Set db = CurrentDb
    Set rs = db.OpenRecordset("FuelSupport_FuelQuotationsT", dbOpenDynaset)

    If Not rs.EOF Then
        rs.MoveFirst ' Move To the first record in the recordset

        ' Loop through the recordset
        Do Until rs.EOF
            If rs("Quote_RefNo").value = currentQuoteRefNo And rs("Airport").value = selectedAirport And rs("FlightType").value = Me.FlightType.value Then
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








