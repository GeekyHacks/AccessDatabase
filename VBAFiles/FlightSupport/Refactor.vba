
Option Compare Database
Option Explicit

'=======================================================
' Event Handlers
'=======================================================

' Ensure a country is selected before focusing on the city field
Private Sub City_GotFocus()
    If IsNull(Me.Country.Value) Then
        MsgBox "Select a country first", vbExclamation
        Me.Country.SetFocus
    End If
End Sub

' Close the form and undo any unsaved changes
Private Sub CloseBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "AddVendorF_V13"
End Sub
Function SanitizeSQL(sql As String, ParamArray params() As Variant) As String
    Dim i As Integer
    Dim paramValue As Variant

    ' Replace each placeholder with the sanitized parameter value
    For i = LBound(params) To UBound(params)
        paramValue = params(i)

        ' Sanitize the parameter value (e.g., escape single quotes)
        If IsNull(paramValue) Then
            paramValue = "NULL"
        ElseIf IsNumeric(paramValue) Then
            ' Numeric values don't need escaping
        Else
            ' Escape single quotes for strings
            paramValue = "'" & Replace(paramValue, "'", "''") & "'"
        End If

        ' Replace the placeholder with the sanitized value
        sql = Replace(sql, "[param" & (i + 1) & "]", paramValue)
    Next i

    ' Return the sanitized SQL query
    SanitizeSQL = sql
End Function
' Update the city dropdown and nationality when a country is selected
Private Sub Country_AfterUpdate()
    On Error GoTo ErrorHandler

    Dim rst As DAO.Recordset
    Dim Nationality As DAO.Recordset
    Dim selectedCountry As Variant
    Dim nationalityValue As String
    Dim sql As String

    ' Get the selected country code
    selectedCountry = Me.Country.Column(0)

    ' Sanitize and build the SQL query for nationality
    sql = SanitizeSQL("SELECT nationality FROM CountriesT_V13 WHERE num_code = [param1];", selectedCountry)
    Set Nationality = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

    ' Check if the recordset is not empty
    If Not Nationality.EOF Then
        nationalityValue = Nationality!Nationality
    Else
        nationalityValue = "" ' Default value
    End If

    ' Close the nationality recordset
    Nationality.Close
    Set Nationality = Nothing

    ' Sanitize and build the SQL query for cities
    sql = SanitizeSQL("SELECT City FROM CitiesT_V13 WHERE CountryCode = [param1];", selectedCountry)
    Me.City.RowSource = sql
    Me.City.Requery

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    If Not Nationality Is Nothing Then
        Nationality.Close
        Set Nationality = Nothing
    End If
End Sub

' Handle date selection for the DateSigned field
Private Sub DateSigned_Click()
    InputDateField DateSigned, "Select a date to use this on your form"
End Sub

' Handle date selection for the AOCExpiry field
Private Sub AOCExpiry_Click()
    InputDateField AOCExpiry, "Select a date to use this on your form"
End Sub

'=======================================================
' Main Function: Add Party Button Click
'=======================================================
Private Sub AddPartyBtn_Click()
    On Error GoTo ErrorHandler

    Dim ws As DAO.Workspace
    Set ws = DBEngine.Workspaces(0) ' Get the default workspace

    Dim PartyID As Long
    Dim locationID As Long
    Dim success As Variant


    ' Log the start of the process
    LogError "Starting AddPartyBtn_Click process.", "AddPartyBtn_Click"

    ' Start the transaction
    ws.BeginTrans
    LogError "Transaction started.", "AddPartyBtn_Click"

    ' Add Party
    PartyID = AddParty()
    If PartyID = -1 Then GoTo Rollback
    LogError "AddParty completed. PartyID: " & PartyID, "AddPartyBtn_Click"

    ' Add Primary Location (from Country, City, and Address controls)
    locationID = AddLocation(PartyID)
    If locationID = -1 Then GoTo Rollback
    LogError "AddLocation completed. LocationID: " & locationID, "AddPartyBtn_Click"

    ' Link Primary Location to Party
    LogError "Before AddPartyLocation", "AddPartyBtn_Click"
    success = ExecuteInTransaction("AddPartyLocation", PartyID, locationID, True) ' True = IsPrimary
    LogError "After AddPartyLocation", "AddPartyBtn_Click"
    If success = -1 Then GoTo Rollback
    LogError "Primary location linked to party.", "AddPartyBtn_Click"
    
' Add Additional Countries (from CountriesListBx)
If Me.MultipleCountriesCbx.Value = -1 Then ' Check if multiple countries are selected
    Dim CountryCode As Long ' Declare the variable here
    Dim i As Variant
    For i = 0 To Me.CountriesListBx.ListCount - 1
        If Me.CountriesListBx.Selected(i) Then
            CountryCode = Me.CountriesListBx.Column(0, i) ' Assign the value here

            ' Link Additional Country to Vendor
            LogError "Before AddVendorCountry", "AddPartyBtn_Click"
            success = ExecuteInTransaction("AddVendorCountry", PartyID, CountryCode)
            LogError "After AddVendorCountry", "AddPartyBtn_Click"
            If success = -1 Then GoTo Rollback
            LogError "Additional country linked to vendor: " & CountryCode, "AddPartyBtn_Click"
        End If
    Next i
End If
    ' Add Party Contact
    Debug.Print "Before AddPartyContact"
    success = ExecuteInTransaction("AddPartyContact", PartyID, locationID)
    Debug.Print "After AddPartyContact"
    If success = -1 Then GoTo Rollback
    LogError "AddPartyContact completed.", "AddPartyBtn_Click"
    
    ' Add Vendor or Client
    If Me.ClientCbx.Value = True Then
        Debug.Print "Before AddPartyDetails (Client)"
        success = ExecuteInTransaction("AddPartyDetails", PartyID, "ClientsT_V13", True) ' Add Client
        Debug.Print "After AddPartyDetails (Client)"
        If success = -1 Then GoTo Rollback
        Debug.Print "Before AddPartyDetails (Vendor)"
        success = ExecuteInTransaction("AddPartyDetails", PartyID, "VendorsT_V13", False) ' Add Vendor
        Debug.Print "After AddPartyDetails (Vendor)"
        If success = -1 Then GoTo Rollback
    Else
        Debug.Print "Before AddPartyDetails (Vendor)"
        success = ExecuteInTransaction("AddPartyDetails", PartyID, "VendorsT_V13", False) ' Add Vendor
        Debug.Print "After AddPartyDetails (Vendor)"
        If success = -1 Then GoTo Rollback
    End If
    LogError "AddPartyDetails completed.", "AddPartyBtn_Click"
    
    ' Add Roles (from ListBox)
    Dim RoleID As Long
    Dim y As Variant
    For Each y In Me.RolesListBox.ItemsSelected
        RoleID = Me.RolesListBox.Column(0, y) ' Assuming the Role ID is in the first column
        Debug.Print "Before AddPartyRole"
        success = ExecuteInTransaction("AddPartyRole", PartyID, RoleID)
        Debug.Print "After AddPartyRole"
        If success = -1 Then GoTo Rollback
        LogError "AddPartyRole completed for RoleID: " & RoleID, "AddPartyBtn_Click"
    Next y

    ' Commit the transaction if everything succeeds
    ws.CommitTrans
    LogError "Transaction committed successfully.", "AddPartyBtn_Click"
    MsgBox "Vendor added successfully!", vbInformation
    Exit Sub

Rollback:
    ' Rollback the transaction on error
    ws.Rollback
    LogError "Transaction rolled back due to an error.", "AddPartyBtn_Click"
    MsgBox "Transaction rolled back due to an error.", vbExclamation
    Exit Sub

ErrorHandler:
    ' Log the error and rollback the transaction
    LogError "Error in AddPartyBtn_Click: " & Err.Description & " (Error Number: " & Err.Number & ")", "AddPartyBtn_Click"
    MsgBox "Error in AddPartyBtn_Click: " & Err.Description, vbCritical
    If Not ws Is Nothing Then
        ws.Rollback ' Ensure the transaction is rolled back
        LogError "Transaction rolled back due to an error.", "AddPartyBtn_Click"
    End If
End Sub

'=======================================================
' Helper Functions
'=======================================================

' Execute an operation within a transaction
Private Function ExecuteInTransaction(operation As String, Optional param1 As Long, Optional param2 As Long, Optional param3 As String, Optional param4 As Boolean) As Variant
    On Error GoTo ErrorHandler

    ' Log the start of the operation
    LogError "Starting operation: " & operation, "ExecuteInTransaction"

    ' Execute the operation based on the provided name
    Select Case operation
        Case "AddParty"
            ExecuteInTransaction = AddParty()
            LogError "AddParty completed. PartyID: " & ExecuteInTransaction, "ExecuteInTransaction"

        Case "AddLocation"
            ExecuteInTransaction = AddLocation(param1)
            LogError "AddLocation completed. LocationID: " & ExecuteInTransaction, "ExecuteInTransaction"

        Case "AddPartyContact"
            ' Validate inputs
            If Not ValidatePartyContact(param1, param2) Then
                ExecuteInTransaction = -1
                Exit Function
            End If

            ' Check for duplicates
            If IsDuplicatePartyContact(param1, param2) Then
                LogError "Duplicate PartyContact entry: PartyID=" & param1 & ", LocationID=" & param2, "ExecuteInTransaction"
                ExecuteInTransaction = -1
                Exit Function
            End If

            ' Add the party contact
            AddPartyContact param1, param2
            ExecuteInTransaction = True
            LogError "AddPartyContact completed.", "ExecuteInTransaction"

        Case "AddPartyLocation"
            ' Validate inputs
            If Not ValidatePartyLocation(param1, param2) Then
                ExecuteInTransaction = -1
                Exit Function
            End If

            ' Check for duplicates
            If IsDuplicatePartyLocation(param1, param2) Then
                LogError "Duplicate PartyLocation entry: PartyID=" & param1 & ", LocationID=" & param2, "ExecuteInTransaction"
                ExecuteInTransaction = -1
                Exit Function
            End If

            ' Add the party location
            AddPartyLocation param1, param2, param4 ' Pass IsPrimary
            ExecuteInTransaction = True
            LogError "AddPartyLocation completed.", "ExecuteInTransaction"

        Case "AddPartyRole"
            ' Add the party role
            AddPartyRole param1, param2
            ExecuteInTransaction = True
            LogError "AddPartyRole completed.", "ExecuteInTransaction"

        Case "AddVendorCountry"
            ' Add the vendor country
            AddVendorCountry param1, param2
            ExecuteInTransaction = True
            LogError "AddVendorCountry completed.", "ExecuteInTransaction"

        Case "AddPartyDetails"
            ' Add party details
            AddPartyDetails param1, param3, param4
            ExecuteInTransaction = True
            LogError "AddPartyDetails completed.", "ExecuteInTransaction"

        Case Else
            LogError "Invalid operation: " & operation, "ExecuteInTransaction"
            MsgBox "Invalid operation: " & operation, vbExclamation
            ExecuteInTransaction = -1
    End Select

    Exit Function

ErrorHandler:
    ' Log the error
    LogError "Error in ExecuteInTransaction: " & Err.Description, "ExecuteInTransaction"
    MsgBox "Error in ExecuteInTransaction: " & Err.Description, vbCritical
    ExecuteInTransaction = -1
End Function

' Generate a unique PartyCode
Private Function GeneratePartyCode() As Long
    On Error GoTo ErrorHandler

    Dim rst As DAO.Recordset
    Dim newCorpID As Long

    ' Step 1: Retrieve the last CorpID from the PartyCodeT_V13 table
    Set rst = CurrentDb.OpenRecordset("PartyCodeT_V13", dbOpenDynaset)

    If Not rst.EOF And Not IsNull(rst!CorpID) Then
        newCorpID = rst!CorpID + 1 ' Increment the last CorpID
    Else
        newCorpID = 1 ' Initialize to 1 if no records exist
    End If

    ' Step 2: Update the PartyCodeT_V13 table with the new CorpID
    If rst.recordCount > 0 Then
        rst.Edit ' Update the existing record
    Else
        rst.AddNew ' Add a new record if the table is empty
    End If
    rst!CorpID = newCorpID
    rst.Update

    rst.Close
    Set rst = Nothing

    ' Return the new CorpID
    GeneratePartyCode = newCorpID
    Exit Function

ErrorHandler:
    LogError "Error in GeneratePartyCode: " & Err.Description, "GeneratePartyCode"
    MsgBox "Error in GeneratePartyCode: " & Err.Description, vbCritical
    GeneratePartyCode = -1
End Function

' Add a new party to the PartiesT_V13 table
Private Function AddParty() As Long
    On Error GoTo ErrorHandler

    ' Validate required fields
    If IsNull(Me.PartyName.Value) Or IsNull(Me.FSANumber.Value) Then
        MsgBox "Party Name and FSANumber are required.", vbExclamation
        AddParty = -1
        Exit Function
    End If

    Dim rst As DAO.Recordset
    Dim CorpID As Long
    Dim PartyCode As String

    ' Step 1: Generate the next CorpID
    CorpID = GeneratePartyCode()
    If CorpID = -1 Then Exit Function ' Exit if there was an error

    ' Step 2: Generate the PartyCode dynamically
    If Me.ClientCbx.Value = True Then
        PartyCode = "CV" & Format(CorpID, "0000") & Format(date, "yy") ' Format as VCXXXXYY
    Else
        PartyCode = "V" & Format(CorpID, "0000") & Format(date, "yy") ' Format as VENXXXXYY
    End If

    ' Step 3: Add the new party to the PartiesT_V13 table
    Set rst = CurrentDb.OpenRecordset("PartiesT_V13", dbOpenDynaset)

    With rst
        .AddNew
        !CorpID = CorpID
        !PartyCode = PartyCode
        !PartyName = Me.PartyName.Value
        !FSANumber = Me.FSANumber.Value
        !AgreementStatus = Me.AgreementStatus.Value
        !AgreementExpiry = Me.AgreementExpiry.Value ' One year date from today
        !DateSigned = Me.DateSigned.Value
        !VAT_NO = Me.VAT_NO.Value
        .Update

        ' Retrieve the newly added CorpID
        .Bookmark = .LastModified
        AddParty = !CorpID
    End With

    rst.Close
    Set rst = Nothing

    ' Display the PartyCode on the form
    DisplayPartyCode CorpID

    Exit Function

ErrorHandler:
    LogError "Error in AddParty: " & Err.Description, "AddParty"
    MsgBox "Error in AddParty: " & Err.Description, vbCritical
    AddParty = -1
End Function

' Generate the PartyCode for display purposes
Private Function GenerateDisplayPartyCode(CorpID As Long) As String
    Dim prefix As String

    ' Determine the prefix based on ClientCbx
    If Me.ClientCbx.Value = True Then
        prefix = "VC" ' Use "VC" for clients that are also vendors
    Else
        prefix = "VEN" ' Use "VEN" for vendors
    End If

    ' Generate the PartyCode in the format [Prefix][XXXX][YY]
    GenerateDisplayPartyCode = prefix & Format(CorpID, "0000") & Format(date, "yy")
End Function

' Display the dynamically generated PartyCode on the form
Private Sub DisplayPartyCode(CorpID As Long)
    Me.PartyCode.Value = GenerateDisplayPartyCode(CorpID)
End Sub

' Add a location to the LocationsT_V13 table
Private Function AddLocation(PartyID As Long) As Long
    On Error GoTo ErrorHandler

    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("LocationsT_V13", dbOpenDynaset)

    With rst
        .AddNew
        !CorpID = PartyID ' Use the passed PartyID
        !Country = Me.Country.Value
        !City = Me.City.Value
        !Address = Me.Address.Value
        .Update

        ' Retrieve the newly added Location ID
        .Bookmark = .LastModified
        AddLocation = !ID
    End With

    rst.Close
    Set rst = Nothing
    Exit Function

ErrorHandler:
    LogError "Error in AddLocation: " & Err.Description, "AddLocation"
    MsgBox "Error in AddLocation: " & Err.Description, vbCritical
    AddLocation = -1
End Function

' Link a vendor to additional countries in VendorCountries_JT_V13
Private Function AddVendorCountry(PartyID As Long, CountryCode As Long) As Boolean
    On Error GoTo ErrorHandler

    ' Validate inputs
    If PartyID <= 0 Or CountryCode <= 0 Then
        MsgBox "Invalid PartyID or CountryCode.", vbExclamation
        AddVendorCountry = False
        Exit Function
    End If

    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("VendorCountries_JT_V13", dbOpenDynaset)

    With rst
        .AddNew
        !CorpID = PartyID
        !CountryCode = CountryCode
        .Update
    End With

    rst.Close
    Set rst = Nothing
    AddVendorCountry = True
    Exit Function

ErrorHandler:
    LogError "Error in AddVendorCountry: " & Err.Description, "AddVendorCountry"
    MsgBox "Error in AddVendorCountry: " & Err.Description, vbCritical
    AddVendorCountry = False
End Function

' Add party contact details to PartyContactT_V13
Private Sub AddPartyContact(PartyID As Long, locationID As Long)
    On Error GoTo ErrorHandler

    ' Log the input values
    LogError "AddPartyContact - PartyID: " & PartyID & ", LocationID: " & locationID, "AddPartyContact"

    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("PartyContactT_V13", dbOpenDynaset)

    With rst
        .AddNew
        !CorpID = PartyID
        !locationID = locationID
        !PrimaryEmail = Me.PrimaryEmail.Value
        !SecondaryEmail = Me.SecondaryEmail.Value
        !Phone = Me.Phone.Value
        .Update
    End With

    rst.Close
    Set rst = Nothing
    Exit Sub

ErrorHandler:
    LogError "Error in AddPartyContact: " & Err.Description & " (PartyID: " & PartyID & ", LocationID: " & locationID & ")", "AddPartyContact"
    MsgBox "Error in AddPartyContact: " & Err.Description, vbCritical
End Sub
Private Function ValidatePartyContact(PartyID As Long, locationID As Long) As Boolean
    Dim rstParty As DAO.Recordset
    Dim rstLocation As DAO.Recordset

    ' Check if PartyID exists in PartiesT_V13
    Set rstParty = CurrentDb.OpenRecordset("SELECT corpID FROM PartiesT_V13 WHERE corpID = " & PartyID, dbOpenSnapshot)
    If rstParty.EOF Then
        MsgBox "Invalid PartyID: " & PartyID, vbExclamation
        ValidatePartyContact = False
        Exit Function
    End If

    ' Check if locationID exists in LocationsT_V13
    Set rstLocation = CurrentDb.OpenRecordset("SELECT ID FROM LocationsT_V13 WHERE ID = " & locationID, dbOpenSnapshot)
    If rstLocation.EOF Then
        MsgBox "Invalid LocationID: " & locationID, vbExclamation
        ValidatePartyContact = False
        Exit Function
    End If

    ValidatePartyContact = True
End Function
Private Function IsDuplicatePartyContact(PartyID As Long, locationID As Long) As Boolean
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("SELECT corpID, locationID FROM PartyContactT_V13 WHERE corpID = " & PartyID & " AND locationID = " & locationID, dbOpenSnapshot)
    IsDuplicatePartyContact = Not rst.EOF
    rst.Close
    Set rst = Nothing
End Function
' Add party details to either ClientsT_V13 or VendorsT_V13
Private Sub AddPartyDetails(PartyID As Long, tableName As String, isClient As Boolean)
    On Error GoTo ErrorHandler

    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset(tableName, dbOpenDynaset)

    With rst
        .AddNew
        !CorpID = PartyID
        !CreditLimit = Me.CreditLimit.Value
        !Deposit = Me.Deposit.Value
        !Currency = Me.Currency.Value
        If isClient Then
            !ClientRating = Me.ClientRating.Value
            !ClientDescription = Me.PartyDescription.Value
        Else
            !VendorRating = Me.VendorRating.Value
            !VendorDescription = Me.PartyDescription.Value
        End If
        .Update
    End With

    rst.Close
    Set rst = Nothing
    Exit Sub

ErrorHandler:
    LogError "Error in AddPartyDetails: " & Err.Description, "AddPartyDetails"
    MsgBox "Error in AddPartyDetails: " & Err.Description, vbCritical
End Sub
' Link a party to a location in PartyLocations_JT_V13
Private Sub AddPartyLocation(PartyID As Long, locationID As Long, IsPrimary As Boolean)
    On Error GoTo ErrorHandler

    ' Log the input values
    LogError "AddPartyLocation - PartyID: " & PartyID & ", LocationID: " & locationID & ", IsPrimary: " & IsPrimary, "AddPartyLocation"

    ' Validate inputs
    If PartyID <= 0 Or locationID <= 0 Then
        MsgBox "Invalid PartyID or LocationID.", vbExclamation
        Exit Sub
    End If

    ' Check for duplicates
    If IsDuplicatePartyLocation(PartyID, locationID) Then
        MsgBox "Duplicate PartyLocation entry: PartyID=" & PartyID & ", LocationID=" & locationID, vbExclamation
        Exit Sub
    End If

    ' Add the party location
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("PartyLocations_JT_V13", dbOpenDynaset)

    With rst
        .AddNew
        !CorpID = PartyID
        !locationID = locationID
        !IsPrimary = IsPrimary
        .Update
    End With

    rst.Close
    Set rst = Nothing
    Exit Sub

ErrorHandler:
    LogError "Error in AddPartyLocation: " & Err.Description & " (PartyID: " & PartyID & ", LocationID: " & locationID & ")", "AddPartyLocation"
    MsgBox "Error in AddPartyLocation: " & Err.Description, vbCritical
End Sub

' Validate PartyID and locationID
Private Function ValidatePartyLocation(PartyID As Long, locationID As Long) As Boolean
    Dim rstParty As DAO.Recordset
    Dim rstLocation As DAO.Recordset

    ' Check if PartyID exists in PartiesT_V13
    Set rstParty = CurrentDb.OpenRecordset("SELECT CorpID FROM PartiesT_V13 WHERE CorpID = " & PartyID, dbOpenSnapshot)
    If rstParty.EOF Then
        MsgBox "Invalid PartyID: " & PartyID, vbExclamation
        ValidatePartyLocation = False
        Exit Function
    End If

    ' Check if locationID exists in LocationsT_V13
    Set rstLocation = CurrentDb.OpenRecordset("SELECT ID FROM LocationsT_V13 WHERE ID = " & locationID, dbOpenSnapshot)
    If rstLocation.EOF Then
        MsgBox "Invalid LocationID: " & locationID, vbExclamation
        ValidatePartyLocation = False
        Exit Function
    End If

    ValidatePartyLocation = True
End Function

' Check for duplicate PartyLocation entries
Private Function IsDuplicatePartyLocation(PartyID As Long, locationID As Long) As Boolean
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("SELECT CorpID, locationID FROM PartyLocations_JT_V13 WHERE CorpID = " & PartyID & " AND locationID = " & locationID, dbOpenSnapshot)
    IsDuplicatePartyLocation = Not rst.EOF
    rst.Close
    Set rst = Nothing
End Function

' Link a party to a role in PartyRolesJT_V13
Private Function AddPartyRole(PartyID As Long, RoleID As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("PartyRolesJT_V13", dbOpenDynaset)

    With rst
        .AddNew
        !CorpID = PartyID ' Use the passed PartyID
        !RoleID = RoleID
        !IsActive = True
        .Update
    End With

    rst.Close
    Set rst = Nothing
    AddPartyRole = True
    Exit Function

ErrorHandler:
    LogError "Error in AddPartyRole: " & Err.Description, "AddPartyRole"
    MsgBox "Error in AddPartyRole: " & Err.Description, vbCritical
    AddPartyRole = False
End Function

'' Log errors to a text file
Private Sub LogError(ErrorMessage As String, ErrorSource As String)
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim fileNumber As Integer
    Dim logMessage As String

    ' Define the file path
    filePath = CurrentProject.Path & "\DebugLog_AddClientF_V13.txt"

    ' Open the file for appending (creates the file if it doesn't exist)
    fileNumber = FreeFile
    Open filePath For Append As #fileNumber

    ' Prepare the log message
    logMessage = "Error Date: " & Now() & vbCrLf & _
                 "Error Source: " & ErrorSource & vbCrLf & _
                 "Error Message: " & ErrorMessage & vbCrLf & _
                 "----------------------------------------" & vbCrLf

    ' Write the log message to the file
    Print #fileNumber, logMessage

    ' Close the file
    Close #fileNumber

    Exit Sub

ErrorHandler:
    MsgBox "Error in LogError: " & Err.Description, vbCritical
End Sub

' Reset the form fields to their default values
Private Sub ResetForm()
    On Error Resume Next ' Skip errors for controls without a Value property

    ' Clear text boxes, combo boxes, and other input fields
    Me.PartyCode.Value = Null
    Me.PartyName.Value = Null
    Me.FSANumber.Value = Null
    Me.AgreementStatus.Value = "Pending" ' "Pending" is the default value
    Me.AgreementExpiry.Value = Null
    Me.DateSigned.Value = Null
    Me.VAT_NO.Value = Null
    Me.Country.Value = Null
    Me.City.Value = Null
    Me.Address.Value = Null
    Me.PrimaryEmail.Value = Null
    Me.SecondaryEmail.Value = Null
    Me.Phone.Value = Null
    Me.CreditLimit.Value = 15000 ' 15000 is the default value
    Me.Deposit.Value = 0 ' 0 is the default value
    Me.Currency.Value = "USD"  ' "USD" is the default value
    Me.ClientRating.Value = 3 ' 3 is the default value
    Me.PartyDescription.Value = Null
    Me.VendorRating.Value = 3 ' 3 is the default value

    ' Reset checkboxes and toggle buttons to default state
    Me.ClientCbx.Value = False
    Me.IsActive.Value = True ' Default to "Active"

    ' Set focus to the first input field
    Me.PartyName.SetFocus

    On Error GoTo 0 ' Reset error handling
End Sub

' Toggle visibility of the CountriesListBx control
Private Sub MultipleCountriesCbx_Click()
    If Me.MultipleCountriesCbx.Value = -1 Then
        Me.CountriesListBx.Visible = True
        Me.CountriesListL.Visible = True
    Else
        Me.CountriesListBx.Visible = False
        Me.CountriesListL.Visible = False
    End If
End Sub

' Reset the form after user confirmation
Private Sub ResetBtn_Click()
    ' Ask the user for confirmation before resetting the form
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to reset the form? All unsaved data will be lost.", vbYesNo + vbQuestion, "Confirm Reset")

    ' Check the user's response
    If response = vbYes Then
        ResetForm ' Reset the form if the user clicks Yes
        LogError "Form reset by user.", "ResetBtn_Click" ' Optional logging
    Else
        ' Do nothing if the user clicks No
        Exit Sub
    End If
End Sub

' Open the UpdateOperatorF form and close the current form
Private Sub Command192_Click()
    DoCmd.OpenForm "UpdateOperatorF", WindowMode:=acWindowNormal
    CloseBtn_Click
End Sub

' Initialize the form when it loads
Private Sub Form_Load()
    Dim fw As New clFormWindow
    DoCmd.GoToRecord , , acNewRec
    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
    Me.ClientRating.Visible = False
End Sub

' Toggle visibility of the ClientRating control
Private Sub ClientCbx_Click()
    If Me.ClientCbx.Value = -1 Then
        Me.ClientRating.Visible = True
    Else
        Me.ClientRating.Visible = False
    End If
End Sub