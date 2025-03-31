Option Compare Database
Option Explicit

'=======================================================
' Enhanced Logging System
'=======================================================
Private Sub LogError( _
    Byval ErrorMessage As String, _
    Byval ErrorSource As String, _
    Optional Byval ErrorNumber As Long = 0, _
    Optional Byval ModuleName As String = "", _
    Optional Byval ProcedureName As String = "", _
    Optional Byval AdditionalInfo As String = "" _
    )
    On Error Goto LogError_Handler
        Dim filePath As String
        Dim fileNumber As Integer
        Dim logMessage As String

        logMessage = String(50, "=") & vbCrLf & _
        "Timestamp: " & Now() & vbCrLf & _
        "Module: " & ModuleName & vbCrLf & _
        "Procedure: " & ProcedureName & vbCrLf & _
        "Error #: " & ErrorNumber & vbCrLf & _
        "Source: " & ErrorSource & vbCrLf & _
        "Message: " & ErrorMessage & vbCrLf & _
        "Additional Info: " & AdditionalInfo & vbCrLf & _
        String(50, "=") & vbCrLf & vbCrLf

        filePath = CurrentProject.Path & "\TransactionAudit.log"
        fileNumber = FreeFile
        Open filePath For Append As #fileNumber
        Print #fileNumber, logMessage
        Close #fileNumber
     Exit Sub

 LogError_Handler:
        MsgBox "Logging System Failure: " & Err.Description, vbCritical
End Sub
Private Function ValidateRequiredFields() As String
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "ValidateRequiredFields"

        ' List of required fields And their display names
        Dim requiredFields As Collection
        Set requiredFields = GetRequiredFieldsCollection()

        ' Log the number of required fields
        LogError "Number of required fields: " & requiredFields.Count, "INFO", 0, "Validation", PROC_NAME

        ' Check If the collection is empty
        If requiredFields.Count = 0 Then
            LogError "Required fields collection is empty", "CRITICAL", 0, "Validation", PROC_NAME
            ValidateRequiredFields = "Unknown field (validation error)"
         Exit Function
        End If

        ' Loop through the required fields And check If they are null
        Dim field As Variant
        For Each field In requiredFields
            ' Log the field being validated
            LogError "Validating field: " & field(1), "INFO", 0, "Validation", PROC_NAME

            If IsFieldMissing(field) Then
                ' Display the error message And Set focus To the missing field
                HandleMissingField field
                ValidateRequiredFields = field(1)
             Exit Function
            End If
        Next field

        ' Validate email format For PrimaryEmail
        If Not IsNull(Me.PrimaryEmail.value) And Me.PrimaryEmail.value <> "" Then
            If Not IsValidEmail(Me.PrimaryEmail.value) Then
                ValidateRequiredFields = "Primary Email (invalid format)"
                MsgBox "The 'Primary Email' field must be a valid email address.", vbExclamation, "Invalid Email Format"
                Me.PrimaryEmail.SetFocus
             Exit Function
            End If
        End If

        ' Validate CountriesListBx only If MultipleCountriesCbx is checked
        If Me.MultipleCountriesCbx.value = True Then
            If Me.CountriesListBx.ItemsSelected.Count = 0 Then
                ValidateRequiredFields = "Countries List"
                MsgBox "You must Select at least one country in 'Countries List'.", vbExclamation, "Missing Required Selection"
                Me.CountriesListBx.SetFocus
             Exit Function
            End If
        End If

        ' If all fields are valid, return an empty string
        ValidateRequiredFields = ""
     Exit Function

 ErrorHandler:
        LogError "Error in ValidateRequiredFields: " & Err.Description & " | Line: " & Erl, "CRITICAL", Err.Number, "Validation", PROC_NAME
        ValidateRequiredFields = "Unknown field (validation error)"
End Function

Private Function IsValidEmail(email As String) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "IsValidEmail"

        ' Simple regex pattern For email validation
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        regex.IgnoreCase = True
        regex.Global = True
        regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"

        ' Check If the email matches the pattern
        IsValidEmail = regex.test(email)

        LogError "Email validation completed: " & IsValidEmail, "INFO", 0, "Validation", PROC_NAME, "Email: " & email
     Exit Function

 ErrorHandler:
        LogError "Failed To validate email: " & Err.Description, "CRITICAL", Err.Number, "Validation", PROC_NAME
        IsValidEmail = False
End Function
Private Function GetRequiredFieldsCollection() As Collection
    On Error Goto ErrorHandler
        Dim requiredFields As Collection
        Set requiredFields = New Collection

        ' Add fields To the collection (Field Control, Display Name)
        requiredFields.aDD Array(Me.PartyName, "Party Name")
        requiredFields.aDD Array(Me.FSANumber, "FSA Number")
        requiredFields.aDD Array(Me.Country, "Country")
        requiredFields.aDD Array(Me.RolesListBox, "Select the role")
        requiredFields.aDD Array(Me.ServiceCategoryList, "Service Category")
        requiredFields.aDD Array(Me.VendorRating, "Vendor Rating")
        requiredFields.aDD Array(Me.VAT_NO, "VAT NO")
        requiredFields.aDD Array(Me.AgreementStatus, "Agreement Status")
        requiredFields.aDD Array(Me.Address, "Address")
        requiredFields.aDD Array(Me.Deposit, "Deposit")
        requiredFields.aDD Array(Me.CreditLimit, "Credit Limit")
        requiredFields.aDD Array(Me.PrimaryEmail, "Primary Email")
        requiredFields.aDD Array(Me.Phone, "Phone")

        ' Add CountriesListBx only If MultipleCountriesCbx is checked
        If Me.MultipleCountriesCbx.value Then
            requiredFields.aDD Array(Me.CountriesListBx, "Countries List")
        End If

        ' Add RolesListBox only If ClientCbx is checked
        If Me.ClientCbx.value Then
            requiredFields.aDD Array(Me.RolesListBox, "Roles List (at least two roles must be selected)")
        End If

        Set GetRequiredFieldsCollection = requiredFields
     Exit Function

 ErrorHandler:
        LogError "Error in GetRequiredFieldsCollection: " & Err.Description & " | Control: " & Err.Source, "CRITICAL", Err.Number, "Validation", "GetRequiredFieldsCollection"
        Set GetRequiredFieldsCollection = New Collection ' Return empty collection To avoid cascading errors
End Function
Private Function IsFieldMissing(field As Variant) As Boolean
    On Error Goto ErrorHandler
        Dim fieldControl As Control
        Dim fieldValue As Variant

        ' Extract the field control
        Set fieldControl = field(0)

        ' Log the control being checked
        LogError "Checking field: " & field(1) & " | Control Type: " & TypeName(fieldControl), "INFO", 0, "Validation", "IsFieldMissing"

        ' Handle list boxes
        If TypeOf fieldControl Is ListBox Then
            If fieldControl.Name = "RolesListBox" And Me.ClientCbx.value Then
                IsFieldMissing = (fieldControl.ItemsSelected.Count < 2)
            Elseif fieldControl.Name = "CountriesListBx" And Me.MultipleCountriesCbx.value Then
                IsFieldMissing = (fieldControl.ItemsSelected.Count = 0)
            Else
                IsFieldMissing = (fieldControl.ItemsSelected.Count = 0)
            End If
        Elseif TypeOf fieldControl Is CheckBox Then
            IsFieldMissing = IsNull(fieldControl.value)
        Else
            fieldValue = fieldControl.value
            IsFieldMissing = (IsNull(fieldValue) Or fieldValue = "")
        End If

        ' Log the result of the check
        LogError "Field missing status: " & IsFieldMissing, "INFO", 0, "Validation", "IsFieldMissing"
     Exit Function

 ErrorHandler:
        LogError "Error in IsFieldMissing: " & Err.Description & " | Control: " & fieldControl.Name, "CRITICAL", Err.Number, "Validation", "IsFieldMissing"
        IsFieldMissing = True ' Treat errors As missing fields
End Function

Private Sub HandleMissingField(field As Variant)
    On Error Goto ErrorHandler
        Dim fieldControl As Control
        Dim fieldName As String

        ' Extract the field control And its display name
        Set fieldControl = field(0)
        fieldName = field(1)

        ' Display a message To the user
        If TypeOf fieldControl Is ListBox Then
            If fieldControl.Name = "RolesListBox" And Me.ClientCbx.value = True Then
                MsgBox "You must Select at least two roles in '" & fieldName & "'.", vbExclamation, "Missing Required Selection"
            Elseif fieldControl.Name = "CountriesListBx" And Me.MultipleCountriesCbx.value = True Then
                MsgBox "You must Select at least one country in '" & fieldName & "'.", vbExclamation, "Missing Required Selection"
            Else
                MsgBox "You must Select at least one item in '" & fieldName & "'.", vbExclamation, "Missing Required Selection"
            End If
        Elseif TypeOf fieldControl Is CheckBox Then
            MsgBox "The checkbox '" & fieldName & "' must be checked.", vbExclamation, "Missing Required Checkbox"
        Else
            MsgBox "The field '" & fieldName & "' is required. Please fill it in.", vbExclamation, "Missing Required Field"
        End If

        ' Set focus To the missing field
        fieldControl.SetFocus
     Exit Sub

 ErrorHandler:
        LogError "Error in HandleMissingField: " & Err.Description & " | Control: " & fieldControl.Name, "CRITICAL", Err.Number, "Validation", "HandleMissingField"
End Sub
'=======================================================
' Core Transaction Handler
'=======================================================
Private Sub AddPartyBtn_Click()
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyBtn_Click"
        Dim ws As DAO.Workspace
        Dim PartyID As Long, locationID As Long
        Dim success As Boolean, CountryCode As Long
        Dim i As Variant, RoleID As Long, y As Variant, x As Variant, ServiceCategoryID As Long
        Dim missingField As String

        LogError "Process initialization", "INFO", 0, "Form", PROC_NAME

        ' =======================================================
        ' Step 1: Validate required fields before starting the transaction
        ' =======================================================
        missingField = ValidateRequiredFields()
        If missingField <> "" Then
            LogError "Validation failed. Missing Or invalid field: " & missingField, "VALIDATION", 0, "Form", PROC_NAME
         Exit Sub
        End If

        Set ws = DBEngine.Workspaces(0)

        ' Start transaction
        On Error Goto TransactionError
            ws.BeginTrans
            LogError "Transaction started", "INFO", 0, "Form", PROC_NAME

            ' Main workflow
            On Error Goto ErrorHandler
                PartyID = ExecuteInTransaction("AddParty")
                If PartyID = -1 Then Goto Rollback

                    locationID = ExecuteInTransaction("AddLocation", PartyID)
                    If locationID = -1 Then Goto Rollback

                        success = ExecuteInTransaction("AddPartyLocation", PartyID, locationID, True)
                        If Not success Then Goto Rollback

                            ' Additional countries
                            If Me.MultipleCountriesCbx.value Then
                                For i = 0 To Me.CountriesListBx.ListCount - 1
                                    If Me.CountriesListBx.Selected(i) Then
                                        CountryCode = Me.CountriesListBx.Column(0, i)
                                        success = ExecuteInTransaction("AddVendorCountry", PartyID, CountryCode)
                                        If Not success Then Goto Rollback
                                        End If
                                    Next i
                                End If

                                ' Contact information
                                success = ExecuteInTransaction("AddPartyContact", PartyID, locationID)
                                If Not success Then Goto Rollback

                                    ' Business details
                                    If Me.ClientCbx.value Then
                                        success = ExecuteInTransaction("AddPartyDetails", PartyID, "ClientsT_V13", True)
                                        If Not success Then Goto Rollback
                                            success = ExecuteInTransaction("AddPartyDetails", PartyID, "VendorsT_V13", False)
                                        Else
                                            success = ExecuteInTransaction("AddPartyDetails", PartyID, "VendorsT_V13", False)
                                        End If
                                        If Not success Then Goto Rollback

                                            ' Role assignments
                                            For Each y In Me.RolesListBox.ItemsSelected
                                                RoleID = Me.RolesListBox.Column(0, y)
                                                success = ExecuteInTransaction("AddPartyRole", PartyID, RoleID)
                                                If Not success Then Goto Rollback
                                                Next y

                                                ' ServiceCategory assignment
                                                For Each x In Me.ServiceCategoryList.ItemsSelected
                                                    ServiceCategoryID = Me.ServiceCategoryList.Column(0, x)
                                                    success = ExecuteInTransaction("AddPartyServiceCategory", PartyID, ServiceCategoryID)
                                                    If Not success Then Goto Rollback
                                                    Next x

                                                    ' Final commit
                                                    ws.CommitTrans
                                                    LogError "Transaction completed successfully", "SUCCESS", 0, "Form", PROC_NAME
                                                    MsgBox "Operation completed successfully!", vbInformation
                                                    ResetForm
                                                    Goto Cleanup ' Skip Rollback And ErrorHandler

 TransactionError:
                                                        LogError "Transaction failure", "CRITICAL", Err.Number, "Form", PROC_NAME, "Line: " & Erl
                                                        Goto Rollback

 Rollback:
                                                            ws.Rollback
                                                            LogError "Transaction rolled back", "CRITICAL", 0, "Form", PROC_NAME
                                                            MsgBox "Operation failed. Check audit log.", vbExclamation
                                                            Goto Cleanup ' Skip ErrorHandler

 ErrorHandler:
                                                                LogError Err.Description, "CRITICAL", Err.Number, "Form", PROC_NAME, _
                                                                "PartyID: " & PartyID & " | LocationID: " & locationID
                                                                Goto Rollback

 Cleanup:
                                                                    ' Ensure the Workspace is properly closed
                                                                    If Not ws Is Nothing Then
                                                                        ws.Close
                                                                        Set ws = Nothing
                                                                    End If
                                                                 Exit Sub
End Sub

'=======================================================
' Database Operations
'=======================================================
Private Function ExecuteInTransaction( _
    Byval operation As String, _
    ParamArray params() As Variant _
    ) As Variant
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "ExecuteInTransaction"
        Dim paramList As String
        Dim i As Long
        Dim rst As DAO.Recordset ' Add this line

        ' Build parameter list manually
        paramList = ""
        For i = LBound(params) To UBound(params)
            paramList = paramList & "Param" & i & ": " & params(i) & " | "
        Next i
        If Len(paramList) > 0 Then
            paramList = Left(paramList, Len(paramList) - 3) ' Remove trailing " | "
        End If

        LogError "Operation started", "INFO", 0, "Transaction", PROC_NAME, _
        "Operation: " & operation & " | Params: " & paramList

        Select Case operation
         Case "AddParty"
            ExecuteInTransaction = AddParty()
         Case "AddLocation"
            ExecuteInTransaction = AddLocation(CLng(params(0)))
         Case "AddPartyLocation"
            ExecuteInTransaction = AddPartyLocation(CLng(params(0)), CLng(params(1)), CBool(params(2)))
         Case "AddVendorCountry"
            ExecuteInTransaction = AddVendorCountry(CLng(params(0)), CLng(params(1)))
         Case "AddPartyContact"
            ExecuteInTransaction = AddPartyContact(CLng(params(0)), CLng(params(1)))
         Case "AddPartyDetails"
            ExecuteInTransaction = AddPartyDetails(CLng(params(0)), CStr(params(1)), CBool(params(2)))
         Case "AddPartyRole"
            ExecuteInTransaction = AddPartyRole(CLng(params(0)), CLng(params(1)))
         Case "AddPartyServiceCategory"
            ExecuteInTransaction = AddPartyServiceCategory(CLng(params(0)), CLng(params(1)))
         Case Else
            LogError "Invalid operation", "ERROR", 9001, "Transaction", PROC_NAME
            ExecuteInTransaction = -1
        End Select

        If ExecuteInTransaction <> -1 Then
            LogError "Operation succeeded", "SUCCESS", 0, "Transaction", PROC_NAME
        End If
     Exit Function

 ErrorHandler:
        LogError "Operation failed: " & Err.Description, "CRITICAL", Err.Number, "Transaction", PROC_NAME
        ExecuteInTransaction = -1
End Function
Private Function AddParty() As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddParty"
        Dim rst As DAO.Recordset
        Dim CorpID As Long

        ' Validation
        If IsNull(Me.PartyName) Or IsNull(Me.FSANumber) Then
            LogError "Missing required fields", "VALIDATION", 1001, "Database", PROC_NAME
            AddParty = -1
         Exit Function
        End If

        ' Check For duplicate PartyName
        If PartyNameExists(Me.PartyName) Then
            LogError "Duplicate PartyName found: " & Me.PartyName, "VALIDATION", 1002, "Database", PROC_NAME
            MsgBox "A party With the name '" & Me.PartyName & "' already exists. Please use a unique name.", vbExclamation, "Duplicate Party Name"
            AddParty = -1
         Exit Function
        End If

        ' Generate ID
        CorpID = GeneratePartyCode()
        If CorpID = -1 Then Exit Function

            ' Create record
            Set rst = CurrentDb.OpenRecordset("PartiesT_V13", dbOpenDynaset)
            With rst
                .AddNew
                !CorpID = CorpID
                !PartyCode = GenerateDisplayPartyCode(CorpID)
                !PartyName = Me.PartyName
                .Update
                .Bookmark = .LastModified
                AddParty = !CorpID
            End With

            LogError "Party created: " & AddParty, "SUCCESS", 0, "Database", PROC_NAME

 Cleanup:
            ' Ensure the Recordset is properly closed
            If Not rst Is Nothing Then
                rst.Close
                Set rst = Nothing
            End If
         Exit Function

 ErrorHandler:
            LogError "Failed To create party", "CRITICAL", Err.Number, "Database", PROC_NAME
            AddParty = -1
            Resume Cleanup
End Function
Private Function PartyNameExists(PartyName As String) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "PartyNameExists"
        Dim rst As DAO.Recordset
        Dim sql As String

        ' Sanitize the PartyName To prevent SQL injection
        PartyName = Replace(PartyName, "'", "''")

        ' Build the SQL query
        sql = "Select 1 FROM PartiesT_V13 WHERE PartyName = '" & PartyName & "'"

        ' Execute the query
        Set rst = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        PartyNameExists = Not rst.EOF

        LogError "PartyNameExists check completed: " & PartyNameExists, "INFO", 0, "Database", PROC_NAME

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To check PartyName existence: " & Err.Description, "CRITICAL", Err.Number, "Database", PROC_NAME
        PartyNameExists = False
        Resume Cleanup
End Function
Private Function AddLocation(PartyID As Long) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddLocation"
        Dim rst As DAO.Recordset

        Set rst = CurrentDb.OpenRecordset("LocationsT_V13", dbOpenDynaset)
        With rst
            .AddNew
            !CorpID = PartyID
            !Country = Me.Country
            !City = Me.City
            !Address = Me.Address
            .Update
            .Bookmark = .LastModified
            AddLocation = !ID
        End With

        LogError "Location created: " & AddLocation, "SUCCESS", 0, "Database", PROC_NAME

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To create location", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddLocation = -1
        Resume Cleanup
End Function
Private Function AddPartyLocation(PartyID As Long, locationID As Long, IsPrimary As Boolean) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyLocation"
        Dim rst As DAO.Recordset

        ' Validate inputs
        If PartyID <= 0 Or locationID <= 0 Then
            LogError "Invalid PartyID Or LocationID", "VALIDATION", 2001, "Database", PROC_NAME
            AddPartyLocation = False
         Exit Function
        End If

        ' Check For duplicates
        If IsDuplicatePartyLocation(PartyID, locationID) Then
            LogError "Duplicate PartyLocation entry", "VALIDATION", 2002, "Database", PROC_NAME
            AddPartyLocation = False
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset("PartyLocations_JT_V13", dbOpenDynaset)
        With rst
            .AddNew
            !CorpID = PartyID
            !locationID = locationID
            !IsPrimary = IsPrimary
            .Update
        End With

        LogError "PartyLocation added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyLocation = True

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyLocation", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyLocation = False
        Resume Cleanup
End Function
Private Function AddVendorCountry(PartyID As Long, CountryCode As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddVendorCountry"
        Dim rst As DAO.Recordset

        ' Validate inputs
        If PartyID <= 0 Or CountryCode <= 0 Then
            LogError "Invalid PartyID Or CountryCode", "VALIDATION", 3001, "Database", PROC_NAME
            AddVendorCountry = False
         Exit Function
        End If

        ' Check For duplicates
        If IsDuplicateVendorCountry(PartyID, CountryCode) Then
            LogError "Duplicate VendorCountry entry", "VALIDATION", 3002, "Database", PROC_NAME
            AddVendorCountry = False
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset("VendorCountries_JT_V13", dbOpenDynaset)
        With rst
            .AddNew
            !CorpID = PartyID
            !CountryCode = CountryCode
            .Update
        End With

        LogError "VendorCountry added", "SUCCESS", 0, "Database", PROC_NAME
        AddVendorCountry = True

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add VendorCountry", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddVendorCountry = False
        Resume Cleanup
End Function
Private Function AddPartyContact(PartyID As Long, locationID As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyContact"
        Dim rst As DAO.Recordset

        ' Validate inputs
        If PartyID <= 0 Or locationID <= 0 Then
            LogError "Invalid PartyID Or LocationID", "VALIDATION", 4001, "Database", PROC_NAME
            AddPartyContact = False
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset("PartyContactT_V13", dbOpenDynaset)
        With rst
            .AddNew
            !CorpID = PartyID
            !locationID = locationID
            !PrimaryEmail = Me.PrimaryEmail
            !SecondaryEmail = Me.SecondaryEmail
            !Phone = Me.Phone
            .Update
        End With

        LogError "PartyContact added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyContact = True

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyContact", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyContact = False
        Resume Cleanup
End Function
Private Function AddPartyDetails(PartyID As Long, TableName As String, isClient As Boolean) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyDetails"
        Dim rst As DAO.Recordset
        Dim RatingValue As Integer
        Dim DescriptionText As String

        ' Validate inputs
        If PartyID <= 0 Or TableName = "" Then
            LogError "Invalid PartyID Or TableName", "VALIDATION", 5001, "Database", PROC_NAME
            AddPartyDetails = False
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset(TableName, dbOpenDynaset)
        With rst
            .AddNew
            !CorpID = PartyID
            !CreditLimit = Nz(Me.CreditLimit, 0)
            !Deposit = Nz(Me.Deposit, 0)
            !Currency = Nz(Me.Currency, "USD")
            !FSANumber = Me.FSANumber
            !AgreementStatus = Me.AgreementStatus
            !AgreementExpiry = Me.AgreementExpiry
            !DateSigned = Me.DateSigned


            If isClient Then
                RatingValue = Nz(Me.ClientRating, 3)
                DescriptionText = Nz(Me.PartyDescription, "")
                !ClientRating = RatingValue
                !ClientDescription = DescriptionText
            Else
                RatingValue = Nz(Me.VendorRating, 3)
                DescriptionText = Nz(Me.PartyDescription, "")
                !VAT_NO = Me.VAT_NO
                !VendorRating = RatingValue
                !VendorDescription = DescriptionText
            End If

            .Update
        End With

        LogError "PartyDetails added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyDetails = True

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyDetails", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyDetails = False
        Resume Cleanup
End Function
Private Function AddPartyRole(PartyID As Long, RoleID As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyRole"
        Dim rst As DAO.Recordset

        LogError "Starting AddPartyRole", "INFO", 0, "Database", PROC_NAME, _
        "PartyID: " & PartyID & " | RoleID: " & RoleID

        ' Validate PartyID
        If Not RecordExists("PartiesT_V13", "CorpID", PartyID) Then
            LogError "Invalid PartyID: " & PartyID, "VALIDATION", 6001, "Database", PROC_NAME
            AddPartyRole = False
         Exit Function
        End If

        ' Validate RoleID
        If Not RecordExists("RolesT_V13", "ID", RoleID) Then
            LogError "Invalid RoleID: " & RoleID, "VALIDATION", 6002, "Database", PROC_NAME
            AddPartyRole = False
         Exit Function
        End If

        ' Check For duplicates
        If IsDuplicatePartyRole(PartyID, RoleID) Then
            LogError "Duplicate PartyRole entry", "VALIDATION", 6003, "Database", PROC_NAME
            AddPartyRole = False
         Exit Function
        End If

        ' Add record
        LogError "Opening PartyRolesJT_V13", "INFO", 0, "Database", PROC_NAME
        Set rst = CurrentDb.OpenRecordset("PartyRolesJT_V13", dbOpenDynaset)
        With rst
            LogError "Adding New record", "INFO", 0, "Database", PROC_NAME
            .AddNew
            !CorpID = PartyID
            !RoleID = RoleID
            !IsActive = True
            !AssignmentDate = date
            .Update
        End With

        LogError "PartyRole added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyRole = True

 Cleanup:
        If Not rst Is Nothing Then
            LogError "Closing recordset", "INFO", 0, "Database", PROC_NAME
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyRole: " & Err.Description, "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyRole = False
        Resume Cleanup
End Function
Private Function AddPartyServiceCategory(PartyID As Long, ServiceCategoryID As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyServiceCategory"
        Dim rst As DAO.Recordset

        LogError "Starting AddPartyServiceCategory", "INFO", 0, "Database", PROC_NAME, _
        "PartyID: " & PartyID & " | ServiceCategoryID: " & ServiceCategoryID

        ' Validate PartyID
        If Not RecordExists("PartiesT_V13", "CorpID", PartyID) Then
            LogError "Invalid PartyID: " & PartyID, "VALIDATION", 6001, "Database", PROC_NAME
            AddPartyServiceCategory = False
         Exit Function
        End If

        ' Validate ServiceCategoryID
        If Not RecordExists("ServiceCategoryT_V13", "ID", ServiceCategoryID) Then
            LogError "Invalid ServiceCategoryID: " & ServiceCategoryID, "VALIDATION", 6002, "Database", PROC_NAME
            AddPartyServiceCategory = False
         Exit Function
        End If

        ' Check For duplicates
        If IsDuplicatePartyServiceCategory(PartyID, ServiceCategoryID) Then
            LogError "Duplicate PartyServiceCategory entry", "VALIDATION", 6003, "Database", PROC_NAME
            AddPartyServiceCategory = False
         Exit Function
        End If

        ' Add record
        LogError "Opening PartyServiceCategoriesJT_V13", "INFO", 0, "Database", PROC_NAME
        Set rst = CurrentDb.OpenRecordset("PartyServiceCategoryJT_V13", dbOpenDynaset)
        With rst
            LogError "Adding New record", "INFO", 0, "Database", PROC_NAME
            .AddNew
            !CorpID = PartyID
            !ServiceCategoryID = ServiceCategoryID
            .Update
        End With

        LogError "PartyServiceCategory added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyServiceCategory = True

 Cleanup:
        If Not rst Is Nothing Then
            LogError "Closing recordset", "INFO", 0, "Database", PROC_NAME
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyServiceCategory: " & Err.Description, "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyServiceCategory = False
        Resume Cleanup
End Function
'=======================================================
' Utility Functions
'=======================================================
Private Function RecordExists(TableName As String, fieldName As String, value As Variant) As Boolean
    On Error Goto ErrorHandler
        Dim rst As DAO.Recordset
        Dim sql As String

        sql = "Select 1 FROM " & TableName & " WHERE " & fieldName & " = " & value
        Set rst = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        RecordExists = Not rst.EOF
        rst.Close
     Exit Function

 ErrorHandler:
        LogError "Failed To check record existence: " & Err.Description, "CRITICAL", Err.Number, "Utility", "RecordExists"
        RecordExists = False
End Function

Private Function JoinParams(params As Variant) As String
    Dim result As String
    Dim i As Long

    result = ""
    For i = LBound(params) To UBound(params)
        result = result & "Param" & i & ": " & params(i) & " | "
    Next i

    If Len(result) > 0 Then
        JoinParams = Left(result, Len(result) - 3) ' Remove trailing " | "
    Else
        JoinParams = "No parameters"
    End If
End Function

Private Function GeneratePartyCode() As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "GeneratePartyCode"
        Dim rst As DAO.Recordset
        Dim newCorpID As Long

        Set rst = CurrentDb.OpenRecordset("PartyCodeT_V13", dbOpenDynaset)
        If Not rst.EOF And Not IsNull(rst!CorpID) Then
            newCorpID = rst!CorpID + 1
        Else
            newCorpID = 1
        End If

        If rst.recordCount > 0 Then
            rst.Edit
        Else
            rst.AddNew
        End If
        rst!CorpID = newCorpID
        rst.Update

        GeneratePartyCode = newCorpID
        LogError "Generated PartyCode: " & newCorpID, "SUCCESS", 0, "Utility", PROC_NAME
     Exit Function

 ErrorHandler:
        LogError "Failed To generate PartyCode", "CRITICAL", Err.Number, "Utility", PROC_NAME
        GeneratePartyCode = -1
End Function

Private Function GenerateDisplayPartyCode(CorpID As Long) As String
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "GenerateDisplayPartyCode"
        Dim prefix As String

        prefix = IIf(Me.ClientCbx.value, "VC", "V")
        GenerateDisplayPartyCode = prefix & Format(CorpID, "0000") & Format(date, "yy")

        LogError "Generated DisplayPartyCode: " & GenerateDisplayPartyCode, "SUCCESS", 0, "Utility", PROC_NAME
     Exit Function

 ErrorHandler:
        LogError "Failed To generate DisplayPartyCode", "CRITICAL", Err.Number, "Utility", PROC_NAME
        GenerateDisplayPartyCode = "ERR-" & Format(Now, "yymmddhhmmss")
End Function

Private Function IsDuplicatePartyLocation(PartyID As Long, locationID As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "IsDuplicatePartyLocation"
        Dim rst As DAO.Recordset

        Set rst = CurrentDb.OpenRecordset( _
        "Select CorpID FROM PartyLocations_JT_V13 " & _
        "WHERE CorpID = " & PartyID & " And locationID = " & locationID, _
        dbOpenSnapshot _
        )
        IsDuplicatePartyLocation = Not rst.EOF

        If IsDuplicatePartyLocation Then
            LogError "Duplicate PartyLocation found", "WARNING", 7001, "Validation", PROC_NAME
        Else
            LogError "No duplicate PartyLocation found", "INFO", 0, "Validation", PROC_NAME
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To check duplicate PartyLocation", "CRITICAL", Err.Number, "Validation", PROC_NAME
        IsDuplicatePartyLocation = False
End Function

Private Function IsDuplicateVendorCountry(PartyID As Long, CountryCode As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "IsDuplicateVendorCountry"
        Dim rst As DAO.Recordset

        Set rst = CurrentDb.OpenRecordset( _
        "Select CorpID FROM VendorCountries_JT_V13 " & _
        "WHERE CorpID = " & PartyID & " And CountryCode = " & CountryCode, _
        dbOpenSnapshot _
        )
        IsDuplicateVendorCountry = Not rst.EOF

        If IsDuplicateVendorCountry Then
            LogError "Duplicate VendorCountry found", "WARNING", 8001, "Validation", PROC_NAME
        Else
            LogError "No duplicate VendorCountry found", "INFO", 0, "Validation", PROC_NAME
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To check duplicate VendorCountry", "CRITICAL", Err.Number, "Validation", PROC_NAME
        IsDuplicateVendorCountry = False
End Function

Private Function IsDuplicatePartyRole(PartyID As Long, RoleID As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "IsDuplicatePartyRole"
        Dim rst As DAO.Recordset

        Set rst = CurrentDb.OpenRecordset( _
        "Select CorpID FROM PartyRolesJT_V13 " & _
        "WHERE CorpID = " & PartyID & " And RoleID = " & RoleID, _
        dbOpenSnapshot _
        )
        IsDuplicatePartyRole = Not rst.EOF

        If IsDuplicatePartyRole Then
            LogError "Duplicate PartyRole found", "WARNING", 9001, "Validation", PROC_NAME
        Else
            LogError "No duplicate PartyRole found", "INFO", 0, "Validation", PROC_NAME
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To check duplicate PartyRole", "CRITICAL", Err.Number, "Validation", PROC_NAME
        IsDuplicatePartyRole = False
End Function
Private Function IsDuplicatePartyServiceCategory(PartyID As Long, ServiceCategoryID As Long) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "IsDuplicatePartyServiceCategory"
        Dim rst As DAO.Recordset

        Set rst = CurrentDb.OpenRecordset( _
        "Select CorpID FROM PartyServiceCategoriesJT_V13 " & _
        "WHERE CorpID = " & PartyID & " And ServiceCategoryID = " & ServiceCategoryID, _
        dbOpenSnapshot _
        )
        IsDuplicatePartyServiceCategory = Not rst.EOF

        If IsDuplicatePartyServiceCategory Then
            LogError "Duplicate PartyServiceCategory found", "WARNING", 7001, "Validation", PROC_NAME
        Else
            LogError "No duplicate PartyServiceCategory found", "INFO", 0, "Validation", PROC_NAME
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To check duplicate PartyServiceCategory", "CRITICAL", Err.Number, "Validation", PROC_NAME
        IsDuplicatePartyServiceCategory = False
End Function
Private Sub ResetForm()
    On Error Resume Next ' Skip errors For controls without a Value Property

    ' Clear text boxes, combo boxes, And other input fields
    Me.PartyCode.value = Null
    Me.PartyName.value = Null
    Me.FSANumber.value = Null
    Me.AgreementStatus.value = "Pending" ' "Pending" is the default value
    Me.AgreementExpiry.value = Null
    Me.DateSigned.value = Null
    Me.VAT_NO.value = Null
    Me.Country.value = Null
    Me.City.value = Null
    Me.Address.value = Null
    Me.PrimaryEmail.value = Null
    Me.SecondaryEmail.value = Null
    Me.Phone.value = Null
    Me.CreditLimit.value = 15000 ' 15000 is the default value
    Me.Deposit.value = 0 ' 0 is the default value
    Me.Currency.value = "USD"  ' "USD" is the default value
    Me.ClientRating.value = 3 ' 3 is the default value
    Me.PartyDescription.value = Null
    Me.VendorRating.value = 3 ' 3 is the default value
    Me.CountriesListBx.RowSource = Me.CountriesListBx.RowSource
    Me.ServiceCategoryList.RowSource = Me.ServiceCategoryList.RowSource
    Me.RolesListBox.RowSource = Me.RolesListBox.RowSource
    Me.MultipleCountriesCbx.value = False
    MultipleCountriesCbx_AfterUpdate
    ' Reset checkboxes And toggle buttons To default state
    Me.ClientCbx.value = False
    ClientCbx_AfterUpdate
    Me.IsActive.value = True ' Default To "Active"

    ' Set focus To the first input field
    Me.PartyName.SetFocus

    On Error Goto 0 ' Reset error handling
End Sub

Private Sub AgreementExpiry_GotFocus()
    InputDateField AgreementExpiry, "Select a date To use this on your form"
End Sub
Private Sub Form_Load()
    Me.ClientCbx.value = 0
    Me.MultipleCountriesCbx.value = 0
    ClientCbx_AfterUpdate
    MultipleCountriesCbx_AfterUpdate
End Sub

Private Sub MultipleCountriesCbx_AfterUpdate()
    If Me.MultipleCountriesCbx.value = -1 Then
        Me.CountriesListBx.Visible = True
        Me.CountriesListL.Visible = True
    Else
        Me.CountriesListBx.Visible = False
        Me.CountriesListL.Visible = False
    End If
End Sub
Private Sub ClientCbx_AfterUpdate()
    If Me.ClientCbx.value = -1 Then
        Me.ClientRating.Visible = True
    Else
        Me.ClientRating.Visible = False
    End If
End Sub
Private Sub City_GotFocus()
    If IsNull(Me.Country.value) Then
        MsgBox "Select a country first", vbExclamation
        Me.Country.SetFocus
    End If
End Sub

' Close the form And undo any unsaved changes
Private Sub CloseBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "AddVendorF_V13"

End Sub
Function SanitizeSQL(sql As String, ParamArray params() As Variant) As String
    Dim i As Integer
    Dim paramValue As Variant

    ' Replace each placeholder With the sanitized parameter value
    For i = LBound(params) To UBound(params)
        paramValue = params(i)

        ' Sanitize the parameter value (e.g., escape single quotes)
        If IsNull(paramValue) Then
            paramValue = "NULL"
        Elseif IsNumeric(paramValue) Then
            ' Numeric values don't need escaping
        Else
            ' Escape single quotes For strings
            paramValue = "'" & Replace(paramValue, "'", "''") & "'"
        End If

        ' Replace the placeholder With the sanitized value
        sql = Replace(sql, "[param" & (i + 1) & "]", paramValue)
    Next i

    ' Return the sanitized SQL query
    SanitizeSQL = sql
End Function
' Update the city dropdown And nationality when a country is selected
Private Sub Country_AfterUpdate()
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Dim Nationality As DAO.Recordset
        Dim selectedCountry As Variant
        Dim nationalityValue As String
        Dim sql As String

        ' Get the selected country code
        selectedCountry = Me.Country.Column(0)

        ' Sanitize And build the SQL query For nationality
        sql = SanitizeSQL("Select nationality FROM CountriesT_V13 WHERE num_code = [param1];", selectedCountry)
        Set Nationality = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

        ' Check If the recordset is Not empty
        If Not Nationality.EOF Then
            nationalityValue = Nationality!Nationality
        Else
            nationalityValue = "" ' Default value
        End If

        ' Close the nationality recordset
        Nationality.Close
        Set Nationality = Nothing

        ' Sanitize And build the SQL query For cities
        sql = SanitizeSQL("Select City FROM CitiesT_V13 WHERE CountryCode = [param1];", selectedCountry)
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

' Handle date selection For the DateSigned field
Private Sub DateSigned_GotFocus()
    InputDateField DateSigned, "Select a date To use this on your form"
End Sub

' Handle date selection For the AOCExpiry field
Private Sub AOCExpiry_GotFocus()
    InputDateField AOCExpiry, "Select a date To use this on your form"
End Sub

Private Sub ResetBtn_Click()
    ResetForm
End Sub