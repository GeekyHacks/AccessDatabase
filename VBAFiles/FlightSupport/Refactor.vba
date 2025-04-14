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
        ' Validate email format For PrimaryEmail
        If Not IsNull(Me.VendorPrimaryEmail.value) And Me.VendorPrimaryEmail.value <> "" Then
            If Not IsValidEmail(Me.VendorPrimaryEmail.value) Then
                ValidateRequiredFields = "VendorPrimaryEmail (invalid format)"
                MsgBox "The 'Vendor Primary Email' field must be a valid email address.", vbExclamation, "Invalid Email Format"
                Me.VendorPrimaryEmail.SetFocus
             Exit Function
            End If
        End If
        ' Validate email format For PrimaryEmail
        If Not IsNull(Me.ClientPrimaryEmail.value) And Me.ClientPrimaryEmail.value <> "" Then
            If Not IsValidEmail(Me.ClientPrimaryEmail.value) Then
                ValidateRequiredFields = "ClientPrimaryEmail (invalid format)"
                MsgBox "The 'Client Primary Email' field must be a valid email address.", vbExclamation, "Invalid Email Format"
                Me.ClientPrimaryEmail.SetFocus
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
        requiredFields.aDD Array(Me.VendorPrimaryEmail, "Vendor Primary Email")
        requiredFields.aDD Array(Me.ClientPrimaryEmail, "Client Primary Email")        requiredFields.aDD Array(Me.Phone, "Phone")

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
Private Sub AddPartyBtn_Click()
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyBtn_Click"
        Dim ws As DAO.Workspace
        Dim PartyID As Long, locationID As Long
        Dim CountryCodes() As Long, CountryNames() As String
        Dim i As Variant, selectedCount As Long
        Dim y As Variant
        Dim x As Variant
        Dim g As Variant
        Dim RoleID As Variant
        Dim RoleName As String
        Dim ServiceCategoryID As Variant
        Dim ServiceCategoryName As String

        ' Validate first
        Dim missingField As String
        missingField = ValidateRequiredFields()
        If missingField <> "" Then
            LogError "Validation failed: " & missingField, "VALIDATION", 0, "Form", PROC_NAME
         Exit Sub
        End If

        ' ===== New: Pre-validate countries BEFORE transaction =====
        If Me.MultipleCountriesCbx.value Then
            ' Count selected countries first
            selectedCount = 0
            For i = 0 To Me.CountriesListBx.ListCount - 1
                If Me.CountriesListBx.Selected(i) Then selectedCount = selectedCount + 1
                Next i

                ' Allocate arrays
                ReDim CountryCodes(1 To selectedCount)
                ReDim CountryNames(1 To selectedCount)

                ' Check For duplicates before transaction
                Dim idx As Long: idx = 1
                For i = 0 To Me.CountriesListBx.ListCount - 1
                    If Me.CountriesListBx.Selected(i) Then
                        CountryCodes(idx) = Me.CountriesListBx.Column(0, i)
                        CountryNames(idx) = Me.CountriesListBx.Column(1, i)

                        ' Check For existing record
                        If IsDuplicateVendorCountry(PartyID, CountryCodes(idx)) Then
                            MsgBox "Country " & CountryNames(idx) & " already exists For this party", vbExclamation
                         Exit Sub
                        End If
                        idx = idx + 1
                    End If
                Next i
            End If

            ' ===== START TRANSACTION (only If all pre-checks pass) =====
            Set ws = DBEngine.Workspaces(0)
            ws.BeginTrans
            On Error Goto Rollback

                ' Core operations
                PartyID = ExecuteInTransaction("AddParty")
                If PartyID = -1 Then Err.Raise 1001, , "AddParty failed"

                    locationID = ExecuteInTransaction("AddLocation", PartyID)
                    If locationID = -1 Then Err.Raise 1002, , "AddLocation failed"

                        ' ===== New: Process pre-validated countries =====
                        If Me.MultipleCountriesCbx.value Then
                            For i = LBound(CountryCodes) To UBound(CountryCodes)
                                If ExecuteInTransaction("AddVendorCountry", PartyID, CountryCodes(i), CountryNames(i)) <= 0 Then
                                    Err.Raise 1004, , "AddVendorCountry failed For " & CountryNames(i)
                                End If
                            Next i
                        End If

                        ' Contact information
                        For Each g In Me.RolesListBox.ItemsSelected
                            RoleID = Me.RolesListBox.Column(0, g)
                            If ExecuteInTransaction("AddPartyContact", PartyID, locationID, RoleID) = -1 Then
                                Err.Raise 1005, , "AddPartyContact failed"
                            End If
                        Next g
                        ' Business details
                        If Me.ClientCbx.value Then
                            ' For Clients - add To ClientsT_V13 (active) And VendorsT_V13 (inactive)
                            If ExecuteInTransaction("AddPartyDetails", PartyID, "ClientsT_V13", True) = -1 Then
                                Err.Raise 1006, , "AddPartyDetails (Client) failed"
                            End If
                            If ExecuteInTransaction("AddPartyDetails", PartyID, "VendorsT_V13", False) = -1 Then
                                Err.Raise 1007, , "AddPartyDetails (Vendor) failed"
                            End If
                        Else
                            ' For non-Clients - only add To VendorsT_V13 As inactive
                            If ExecuteInTransaction("AddPartyDetails", PartyID, "VendorsT_V13", False) = -1 Then
                                Err.Raise 1008, , "AddPartyDetails (Vendor) failed"
                            End If
                        End If

                        ' Role assignments With New name parameter
                        For Each y In Me.RolesListBox.ItemsSelected
                            RoleID = Me.RolesListBox.Column(0, y)
                            RoleName = Me.RolesListBox.Column(1, y)
                            If ExecuteInTransaction("AddPartyRole", PartyID, RoleID, RoleName) = -1 Then
                                Err.Raise 1007, , "AddPartyRole failed"
                            End If
                        Next y

                        ' ServiceCategory assignment With New name
                        For Each x In Me.ServiceCategoryList.ItemsSelected
                            ServiceCategoryID = Me.ServiceCategoryList.Column(0, x)
                            ServiceCategoryName = Me.ServiceCategoryList.Column(1, x)
                            If ExecuteInTransaction("AddPartyServiceCategory", PartyID, ServiceCategoryID, ServiceCategoryName) = -1 Then
                                Err.Raise 1008, , "AddPartyServiceCategory failed"
                            End If
                        Next x

                        ' Final commit
                        ws.CommitTrans
                        LogError "Transaction completed successfully", "SUCCESS", 0, "Form", PROC_NAME
                        MsgBox "Operation completed successfully!", vbInformation
                        ResetForm
                        Goto Cleanup

 Rollback:
                            LogError "Transaction failed: " & Err.Description, "CRITICAL", Err.Number, "Form", PROC_NAME
                            If Not ws Is Nothing Then ws.Rollback
                                MsgBox "Operation failed: " & Err.Description & vbCrLf & "No changes were saved.", vbExclamation

 Cleanup:
                                If Not ws Is Nothing Then ws.Close
                                 Exit Sub

 ErrorHandler:
                                    LogError "Unexpected error: " & Err.Description, "CRITICAL", Err.Number, "Form", PROC_NAME
                                    Resume Rollback
End Sub

Private Function ExecuteInTransaction( _
    Byval operation As String, _
    ParamArray params() As Variant _
    ) As Variant
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "ExecuteInTransaction"
        Dim paramList As String
        Dim i As Long

        ' Build parameter list For logging
        paramList = ""
        For i = LBound(params) To UBound(params)
            paramList = paramList & "Param" & i & ": " & params(i) & " | "
        Next i
        If Len(paramList) > 0 Then paramList = Left(paramList, Len(paramList) - 3)

            LogError "Operation started", "INFO", 0, "Transaction", PROC_NAME, _
            "Operation: " & operation & " | Params: " & paramList

            Select Case operation
             Case "AddParty"
                ExecuteInTransaction = AddParty()
             Case "AddLocation"
                ExecuteInTransaction = AddLocation(CLng(params(0)))
             Case "AddVendorCountry"
                ExecuteInTransaction = AddVendorCountry(CLng(params(0)), CLng(params(1)), CStr(params(2)))
             Case "AddPartyContact"
                ExecuteInTransaction = AddPartyContact(CLng(params(0)), CLng(params(1)), CLng(params(2)))
             Case "AddPartyDetails"
                ExecuteInTransaction = AddPartyDetails(CLng(params(0)), CStr(params(1)), CBool(params(2)))
             Case "AddPartyRole"
                ExecuteInTransaction = AddPartyRole(CLng(params(0)), CLng(params(1)), CStr(params(2)))
             Case "AddPartyServiceCategory"
                ExecuteInTransaction = AddPartyServiceCategory(CLng(params(0)), CLng(params(1)), CStr(params(2)))
             Case Else
                LogError "Invalid operation", "ERROR", 9001, "Transaction", PROC_NAME
                ExecuteInTransaction = -1
            End Select

            If ExecuteInTransaction = -1 Then
                LogError "Operation failed", "ERROR", 0, "Transaction", PROC_NAME
            Else
                LogError "Operation succeeded", "SUCCESS", 0, "Transaction", PROC_NAME
            End If
         Exit Function

 ErrorHandler:
            LogError "Operation failed: " & Err.Description & " | Return: " & ExecuteInTransaction, _
            "CRITICAL", Err.Number, "Transaction", PROC_NAME
            ExecuteInTransaction = -1
End Function
Private Function AddParty() As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddParty"
        Dim rst As DAO.Recordset
        Dim corpID As Long
        Dim PartyCode As String

        ' Generate IDs first
        corpID = GeneratePartyCode()
        If corpID <= 0 Then Err.Raise 1001, , "Invalid PartyCode"

            PartyCode = GenerateDisplayPartyCode(corpID)
            If Len(PartyCode) = 0 Then Err.Raise 1002, , "Invalid DisplayPartyCode"

                Set rst = CurrentDb.OpenRecordset("PartiesT_V13", dbOpenDynaset, dbAppendOnly)
                With rst
                    .AddNew
                    !corpID = corpID  ' Ensure correct Case For field name
                    !PartyCode = PartyCode
                    !PartyName = Nz(Me.PartyName.value, "")
                    .Update
                    .Bookmark = .LastModified
                    AddParty = !corpID
                End With

                LogError "Party created: " & AddParty, "SUCCESS", 0, "Database", PROC_NAME
             Exit Function

 ErrorHandler:
                LogError "AddParty failed: " & Err.Description, "CRITICAL", Err.Number, "Database", PROC_NAME
                AddParty = -1
                If Not rst Is Nothing Then
                    If rst.EditMode <> dbEditNone Then rst.CancelUpdate
                        rst.Close
                    End If
End Function
Private Function AddLocation(PartyID As Long) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddLocation"
        Dim rst As DAO.Recordset

        ' Validate input first
        If PartyID <= 0 Then
            LogError "Invalid PartyID", "VALIDATION", 2001, "Database", PROC_NAME
            AddLocation = -1
         Exit Function
        End If

        ' Prepare data before opening recordset
        Dim CountryCode As Long: CountryCode = Me.Country.Column(0)
        Dim CountryName As String: CountryName = Me.Country.Column(1)
        Dim City As String: City = Me.City.value
        Dim Address As String: Address = Me.Address.value

        Set rst = CurrentDb.OpenRecordset("LocationsT_V13", dbOpenDynaset, dbAppendOnly)
        With rst
            .AddNew
            !corpID = PartyID
            !CountryCode = CountryCode
            !CountryName = CountryName
            !City = City
            !Address = Address
            !isPrimary = True
            .Update
            .Bookmark = .LastModified
            AddLocation = !ID
        End With

        LogError "Location created: " & AddLocation, "SUCCESS", 0, "Database", PROC_NAME
     Exit Function

 ErrorHandler:
        LogError "Failed To create location", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddLocation = -1
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Function
Private Function AddVendorCountry(PartyID As Long, CountryCode As Long, CountryName As String) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddVendorCountry"
        Dim rst As DAO.Recordset
        Dim newID As Long

        Set rst = CurrentDb.OpenRecordset("PartyLocations_JT_V13", dbOpenDynaset, dbAppendOnly)
        With rst
            .AddNew
            !corpID = PartyID
            !CountryCode = CountryCode
            !CountryName = CountryName
            .Update
            .Bookmark = .LastModified
            newID = !ID  ' Get auto-generated ID
        End With

        LogError "VendorCountry added", "SUCCESS", 0, "Database", PROC_NAME
        AddVendorCountry = newID
     Exit Function

 ErrorHandler:
        LogError "Failed To add VendorCountry", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddVendorCountry = -1
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
Private Function AddPartyContact(PartyID As Long, locationID As Long, RoleID As Long) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyContact"
        Dim isVendor As Boolean
        Dim isClient As Boolean
        Dim rst As DAO.Recordset
        Dim newID As Long

        ' Validate inputs
        If PartyID <= 0 Or locationID <= 0  Or RoleID <= 0 Then
            LogError "Invalid IDs provided", "VALIDATION", 4001, "Database", PROC_NAME
            AddPartyContact = -1
         Exit Function
        End If
        ' === ROLE DETECTION ===
        isVendor = (RoleID = GetVendorRoleID())
        isClient = (RoleID = GetClientRoleID())

        If RecordExists("PartyContactT_V13", , , "[CorpID] = " & PartyID & " And [Role_ID] = " & RoleID) Then
            LogError "Duplicate contact For RoleID " & RoleID, "VALIDATION", 4003, "Database", PROC_NAME
         Exit Function
        End If


        ' Add record
        Set rst = CurrentDb.OpenRecordset("PartyContactT_V13", dbOpenDynaset, dbAppendOnly)
        With rst
            .AddNew
            !corpID = PartyID
            !locationID = locationID
            !Role_ID = RoleID ' Store

            ' --- Set contact info based on role ---
            Select Case True
             Case isVendor:
                !PrimaryEmail = Nz(Me.VendorPrimaryEmail.value, "")
                !SecondaryEmail = Nz(Me.SecondaryEmail.value, "")
                !Phone = Nz(Me.Phone.value, "")

             Case isClient:
                !PrimaryEmail = Nz(Me.ClientPrimaryEmail.value, "")
                !SecondaryEmail = Nz(Me.SecondaryEmail.value, "")
                !Phone = Nz(Me.Phone.value, "")

            End Select

            .Update
        End With

        LogError "Contact added For " & "RoleID ", _
        "SUCCESS", 0, "Database", PROC_NAME
        AddPartyContact = 0

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add contact For RoleID " & RoleID & Err.Description, _
        "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyContact = -1
        Resume Cleanup
End Function
Private Function GetVendorRoleID() As Long
    ' Retrieve from config table Or constant
    GetVendorRoleID = DLookup("ID", "RolesT_V13", "RoleName = 'Vendor'")
End Function

Private Function GetClientRoleID() As Long
    ' Retrieve from config table Or constant
    GetClientRoleID = DLookup("ID", "RolesT_V13", "RoleName = 'Client'")
End Function
Private Function AddPartyDetails(PartyID As Long, tableName As String, isClient As Boolean) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyDetails"
        Dim rst As DAO.Recordset
        Dim newID As Long
        Dim RatingValue As Integer
        Dim DescriptionText As String

        ' Validate inputs
        If PartyID <= 0 Or tableName = "" Then
            LogError "Invalid PartyID Or TableName", "VALIDATION", 5001, "Database", PROC_NAME
            AddPartyDetails = -1
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset(tableName, dbOpenDynaset, dbAppendOnly)
        With rst
            .AddNew
            !corpID = PartyID
            !CreditLimit = Nz(Me.CreditLimit, 0)
            !Deposit = Nz(Me.Deposit, 0)
            !Currency = Nz(Me.Currency, "USD")
            !FSANumber = Me.FSANumber
            !AgreementStatus = Me.AgreementStatus
            !AgreementExpiry = Me.AgreementExpiry
            !DateSigned = Me.DateSigned
            !IsActive = True

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
            .Bookmark = .LastModified
            newID = !ID ' Get auto-generated ID
        End With

        LogError "PartyDetails added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyDetails = newID

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyDetails", "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyDetails = -1
        Resume Cleanup
End Function
Private Function AddPartyRole(PartyID As Long, RoleID As Long, RoleName As String) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyRole"
        Dim rst As DAO.Recordset
        Dim newID As Long

        ' Validate PartyID
        If Not RecordExists("PartiesT_V13", "CorpID", PartyID) Then
            LogError "Invalid PartyID: " & PartyID, "VALIDATION", 6001, "Database", PROC_NAME
            AddPartyRole = -1
         Exit Function
        End If

        ' Validate RoleID
        If Not RecordExists("RolesT_V13", "ID", RoleID) Then
            LogError "Invalid RoleID: " & RoleID, "VALIDATION", 6002, "Database", PROC_NAME
            AddPartyRole = -1
         Exit Function
        End If

        ' Check For duplicates
        If IsDuplicatePartyRole(PartyID, RoleID) Then
            LogError "Duplicate PartyRole entry", "VALIDATION", 6003, "Database", PROC_NAME
            AddPartyRole = -2 ' Special code For duplicates
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset("PartyRolesJT_V13", dbOpenDynaset, dbAppendOnly)
        With rst
            .AddNew
            !corpID = PartyID
            !RoleID = RoleID
            !RoleName = RoleName
            !AssignmentDate = date
            .Update
            .Bookmark = .LastModified
            newID = !ID
        End With

        LogError "PartyRole added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyRole = newID

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyRole: " & Err.Description, "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyRole = -1
        Resume Cleanup
End Function
Private Function AddPartyServiceCategory(PartyID As Long, ServiceCategoryID As Long, Categoryname As String) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyServiceCategory"
        Dim rst As DAO.Recordset
        Dim newID As Long

        ' Validate PartyID
        If Not RecordExists("PartiesT_V13", "CorpID", PartyID) Then
            LogError "Invalid PartyID: " & PartyID, "VALIDATION", 6001, "Database", PROC_NAME
            AddPartyServiceCategory = -1
         Exit Function
        End If

        ' Validate ServiceCategoryID
        If Not RecordExists("ServiceCategoryT_V13", "ID", ServiceCategoryID) Then
            LogError "Invalid ServiceCategoryID: " & ServiceCategoryID, "VALIDATION", 6002, "Database", PROC_NAME
            AddPartyServiceCategory = -1
         Exit Function
        End If

        ' Check For duplicates
        If IsDuplicatePartyServiceCategory(PartyID, ServiceCategoryID) Then
            LogError "Duplicate PartyServiceCategory entry", "VALIDATION", 6003, "Database", PROC_NAME
            AddPartyServiceCategory = -2 ' Special code For duplicates
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset("PartyServiceCategoryJT_V13", dbOpenDynaset, dbAppendOnly)
        With rst
            .AddNew
            !corpID = PartyID
            !ServiceCategoryID = ServiceCategoryID
            !Categoryname = Categoryname
            .Update
            .Bookmark = .LastModified
            newID = !ID
        End With

        LogError "PartyServiceCategory added", "SUCCESS", 0, "Database", PROC_NAME
        AddPartyServiceCategory = newID

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add PartyServiceCategory: " & Err.Description, "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyServiceCategory = -1
        Resume Cleanup
End Function
'=======================================================
' Utility Functions
'=======================================================
Private Function RecordExists(tableName As String, fieldName As String, value As Variant) As Boolean
    On Error Goto ErrorHandler
        Dim rst As DAO.Recordset
        Dim sql As String

        sql = "Select 1 FROM " & tableName & " WHERE " & fieldName & " = " & value
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
        If Not rst.EOF And Not IsNull(rst!corpID) Then
            newCorpID = rst!corpID + 1
        Else
            newCorpID = 1
        End If

        If rst.recordCount > 0 Then
            rst.Edit
        Else
            rst.AddNew
        End If
        rst!corpID = newCorpID
        rst.Update

        GeneratePartyCode = newCorpID
        LogError "Generated PartyCode: " & newCorpID, "SUCCESS", 0, "Utility", PROC_NAME
     Exit Function

 ErrorHandler:
        LogError "Failed To generate PartyCode", "CRITICAL", Err.Number, "Utility", PROC_NAME
        GeneratePartyCode = -1
End Function

Private Function GenerateDisplayPartyCode(corpID As Long) As String
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "GenerateDisplayPartyCode"
        Dim prefix As String

        prefix = IIf(Me.ClientCbx.value, "VC", "V")
        GenerateDisplayPartyCode = prefix & Format(corpID, "0000") & Format(date, "yy")

        LogError "Generated DisplayPartyCode: " & GenerateDisplayPartyCode, "SUCCESS", 0, "Utility", PROC_NAME
     Exit Function

 ErrorHandler:
        LogError "Failed To generate DisplayPartyCode", "CRITICAL", Err.Number, "Utility", PROC_NAME
        GenerateDisplayPartyCode = "ERR-" & Format(Now, "yymmddhhmmss")
End Function
Private Function IsDuplicateVendorCountry(PartyID As Long, CountryCode As Long) As Boolean
    Const PROC_NAME As String = "IsDuplicateVendorCountry"
    Dim rst As DAO.Recordset
    Dim sql As String

    sql = "Select COUNT(*) FROM PartyLocations_JT_V13 " & _
    "WHERE corpID = " & PartyID & " And CountryCode = " & CountryCode

    Set rst = CurrentDb.OpenRecordset(sql)
    IsDuplicateVendorCountry = (rst.Fields(0).value > 0)
    rst.Close

    LogError "Duplicate check completed", "INFO", 0, "Validation", PROC_NAME, _
    "Duplicate: " & IsDuplicateVendorCountry
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
    Me.ClientPrimaryEmail.value = Null
    Me.VendorPrimaryEmail.value = Null
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
    If Me.ClientCbx.value = 0 Then
        Me.ClientPrimaryEmail.Visible = False
        Me.ClientRating.Visible = False
    Else

        Me.ClientRating.Visible = True
        Me.ClientPrimaryEmail.Visible = True
    End If

    If Me.ClientCbx.value = -1 Then
        MsgBox "Please add the Client Primary email, it might be different"
        Me.ClientPrimaryEmail.Visible = True
        Me.ClientRating.Visible = True
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



