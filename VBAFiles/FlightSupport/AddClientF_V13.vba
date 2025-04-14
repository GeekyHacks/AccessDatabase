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
        If Not IsNull(Me.ClientPrimaryEmail.value) And Me.ClientPrimaryEmail.value <> "" Then
            If Not IsValidEmail(Me.ClientPrimaryEmail.value) Then
                ValidateRequiredFields = "ClientPrimaryEmail (invalid format)"
                MsgBox "The 'Client Primary Email' field must be a valid email address.", vbExclamation, "Invalid Email Format"
                Me.ClientPrimaryEmail.SetFocus
             Exit Function
            End If
        End If
        ' Validate email format For PrimaryEmail
        If Not IsNull(Me.VendorPrimaryEmail.value) And Me.VendorPrimaryEmail.value <> "" Then
            If Not IsValidEmail(Me.VendorPrimaryEmail.value) Then
                ValidateRequiredFields = "VendorPrimaryEmail (invalid format)"
                MsgBox "The 'Vendor Primary Email' field must be a valid email address.", vbExclamation, "Invalid Email Format"
                Me.VendorPrimaryEmail.SetFocus
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
        requiredFields.aDD Array(Me.VendorRating, "Vendor Rating")
        requiredFields.aDD Array(Me.VAT_NO, "VAT NO")
        requiredFields.aDD Array(Me.AgreementStatus, "Agreement Status")
        requiredFields.aDD Array(Me.Address, "Address")
        requiredFields.aDD Array(Me.Deposit, "Deposit")
        requiredFields.aDD Array(Me.CreditLimit, "Credit Limit")
        requiredFields.aDD Array(Me.VendorPrimaryEmail, "Vendor Primary Email")
        requiredFields.aDD Array(Me.ClientPrimaryEmail, "Client Primary Email")
        requiredFields.aDD Array(Me.Phone, "Phone")

        ' Add RolesListBox only If VendorChx is checked
        If Me.VendorChx.value Then
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
            If fieldControl.Name = "RolesListBox" Or Me.VendorChx.value Then
                IsFieldMissing = (fieldControl.ItemsSelected.Count < 2)
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
            If fieldControl.Name = "RolesListBox" And Me.VendorChx.value = True Then
                MsgBox "You must Select at least two roles in '" & fieldName & "'.", vbExclamation, "Missing Required Selection"
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
        Dim success As Boolean
        Dim selectedCount As Long
        Dim RoleName As String
        Dim i As Variant, RoleID As Long, y As Variant, x As Variant, ServiceCategoryID As Long, g As Variant
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
            Dim inTransaction As Boolean
            inTransaction = True

            LogError "Transaction started", "INFO", 0, "Form", PROC_NAME

            ' Main workflow
            On Error Goto ErrorHandler
                PartyID = ExecuteInTransaction("AddParty")
                If PartyID = -1 Then Goto Rollback
                    ' Add party location
                    locationID = ExecuteInTransaction("AddLocation", PartyID) ' Now handles IsPrimary internally
                    If locationID = -1 Then Goto Rollback
                        ' Add party contact
                        For Each g In Me.RolesListBox.ItemsSelected
                            RoleID = Me.RolesListBox.Column(0, g)
                            success = ExecuteInTransaction("AddPartyContact", PartyID, locationID, RoleID)
                            If Not success Then Goto Rollback
                            Next g

                            ' success = ExecuteInTransaction("AddPartyContact", PartyID, locationID)
                            ' If Not success Then Goto Rollback

                            ' Business details (now includes IsActive)
                            If Me.VendorChx.value Then
                                success = ExecuteInTransaction("AddPartyDetails", PartyID, "VendorsT_V13", True)
                                If Not success Then Goto Rollback
                                    success = ExecuteInTransaction("AddPartyDetails", PartyID, "ClientsT_V13", False)
                                Else
                                    success = ExecuteInTransaction("AddPartyDetails", PartyID, "ClientsT_V13", False)
                                End If
                                If Not success Then Goto Rollback

                                    ' Role assignments (simplified - no IsActive)
                                    For Each y In Me.RolesListBox.ItemsSelected
                                        RoleID = Me.RolesListBox.Column(0, y)
                                        RoleName = Me.RolesListBox.Column(1, y)
                                        If ExecuteInTransaction("AddPartyRole", PartyID, RoleID, RoleName) = -1 Then
                                            Err.Raise 1007, , "AddPartyRole failed"
                                        End If
                                    Next y

                                    ' Final commit
                                    ws.CommitTrans
                                    LogError "Transaction completed successfully", "SUCCESS", 0, "Form", PROC_NAME
                                    MsgBox "Operation completed successfully!", vbInformation
                                    ResetForm
                                    Goto Cleanup

 TransactionError:
                                        LogError "Transaction failure", "CRITICAL", Err.Number, "Form", PROC_NAME, "Line: " & Erl
                                        Goto Rollback

 Rollback:
                                            If inTransaction Then
                                                ws.Rollback
                                                LogError "Transaction rolled back", "CRITICAL", 0, "Form", PROC_NAME
                                                MsgBox "Operation failed. Check audit log.", vbExclamation
                                                Goto Cleanup
                                                    inTransaction = False
                                                End If

 ErrorHandler:
                                                LogError Err.Description, "CRITICAL", Err.Number, "Form", PROC_NAME, _
                                                "PartyID: " & PartyID & " | LocationID: " & locationID
                                                Goto Rollback

 Cleanup:
                                                    If Not ws Is Nothing Then
                                                        inTransaction = False
                                                        ws.Close
                                                        Set ws = Nothing
                                                    End If
                                                 Exit Sub
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
        If Len(paramList) > 0 Then
            paramList = Left(paramList, Len(paramList) - 3)
        End If

        LogError "Operation started", "INFO", 0, "Transaction", PROC_NAME, _
        "Operation: " & operation & " | Params: " & paramList

        Select Case operation
         Case "AddParty"
            ExecuteInTransaction = AddParty()
         Case "AddLocation"
            ExecuteInTransaction = AddLocation(CLng(params(0)))
         Case "AddPartyContact"
            ExecuteInTransaction = AddPartyContact(CLng(params(0)), CLng(params(1)), CLng(params(2)))
         Case "AddPartyDetails"
            ' Now includes IsActive in the target tables
            ExecuteInTransaction = AddPartyDetails(CLng(params(0)), CStr(params(1)), CBool(params(2)))
         Case "AddPartyRole"
            ' Simplified - no longer handles IsActive
            ExecuteInTransaction = AddPartyRole(CLng(params(0)), CLng(params(1)), CStr(params(2)))
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
        Dim corpID As Long

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
        corpID = GeneratePartyCode()
        If corpID = -1 Then Exit Function

            ' Create record
            Set rst = CurrentDb.OpenRecordset("PartiesT_V13", dbOpenDynaset)
            With rst
                .AddNew
                !corpID = corpID
                !PartyCode = GenerateDisplayPartyCode(corpID)
                !PartyName = Me.PartyName
                .Update
                .Bookmark = .LastModified
                AddParty = !corpID
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
Private Function AddPartyContact( _
    PartyID As Long, _
    locationID As Long, _
    RoleID As Long _
    ) As Long
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyContact"
        Dim rst As DAO.Recordset
        Dim isVendor As Boolean
        Dim isClient As Boolean

        ' === VALIDATION ===
        If PartyID <= 0 Or locationID <= 0 Or RoleID <= 0 Then
            LogError "Invalid IDs provided", "VALIDATION", 4001, "Database", PROC_NAME
         Exit Function
        End If

        ' === ROLE DETECTION ===
        isVendor = (RoleID = GetVendorRoleID())
        isClient = (RoleID = GetClientRoleID())

        ' Check For existing contact For this role
        If RecordExists("PartyContactT_V13", , , "[CorpID] = " & PartyID & " And [Role_ID] = " & RoleID) Then
            LogError "Duplicate contact For RoleID " & RoleID, "VALIDATION", 4003, "Database", PROC_NAME
         Exit Function
        End If

        ' === DATABASE OPERATION ===
        Set rst = CurrentDb.OpenRecordset("PartyContactT_V13", dbOpenDynaset)

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
        AddPartyContact = True

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add contact For RoleID " & RoleID & Err.Description, _
        "CRITICAL", Err.Number, "Database", PROC_NAME
        AddPartyContact = False
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
Private Function AddPartyDetails(PartyID As Long, tableName As String, isVendor As Boolean) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyDetails"
        Dim rst As DAO.Recordset
        Dim RatingValue As Integer
        Dim DescriptionText As String

        ' Validate inputs
        If PartyID <= 0 Or tableName = "" Then
            LogError "Invalid PartyID Or TableName", "VALIDATION", 5001, "Database", PROC_NAME
            AddPartyDetails = False
         Exit Function
        End If

        ' Add record
        Set rst = CurrentDb.OpenRecordset(tableName, dbOpenDynaset)
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
            !IsActive = True  ' Added IsActive field

            If isVendor Then
                RatingValue = Nz(Me.VendorRating, 3)
                DescriptionText = Nz(Me.PartyDescription, "")
                !VendorRating = RatingValue
                !VendorDescription = DescriptionText
            Else
                RatingValue = Nz(Me.ClientRating, 3)
                DescriptionText = Nz(Me.PartyDescription, "")
                !ClientRating = RatingValue
                !VAT_NO = Me.VAT_NO
                !ClientDescription = DescriptionText
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

'=======================================================
' Utility Functions
'=======================================================
Private Function RecordExists( _
    tableName As String, _
    Optional fieldName As String = "", _
    Optional value As Variant, _
    Optional whereClause As String = "" _
    ) As Boolean
    On Error Goto ErrorHandler
        Dim rst As DAO.Recordset
        Dim sql As String

        ' Build the SQL query With proper field delimiters
        If whereClause <> "" Then
            sql = "Select 1 FROM [" & tableName & "] WHERE " & whereClause
        Elseif fieldName <> "" Then
            sql = "Select 1 FROM [" & tableName & "] WHERE [" & fieldName & "] = " & ToSQLValue(value)
        Else
            sql = "Select 1 FROM [" & tableName & "] WHERE 1=0"
        End If

        ' Debug output (remove after testing)
        Debug.Print "Executing SQL: " & sql

        ' Execute query
        Set rst = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
        RecordExists = Not rst.EOF
        rst.Close

     Exit Function

 ErrorHandler:
        LogError "Failed To check record existence", "CRITICAL", Err.Number, "Utility", "RecordExists", _
        "SQL: " & sql & " | Error: " & Err.Description
        RecordExists = False
End Function
Private Function ToSQLValue(value As Variant) As String
    ' Helper Function To properly format values
    If IsNull(value) Then
        ToSQLValue = "NULL"
    Elseif VarType(value) = vbString Then
        ToSQLValue = "'" & Replace(value, "'", "''") & "'"
    Elseif VarType(value) = vbDate Then
        ToSQLValue = "#" & Format(value, "yyyy-mm-dd") & "#"
    Else
        ToSQLValue = CStr(value)
    End If
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

        prefix = IIf(Me.VendorChx.value, "VC", "C")
        GenerateDisplayPartyCode = prefix & Format(corpID, "000000") & Format(date, "yy")

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
    Me.RolesListBox.RowSource = Me.RolesListBox.RowSource
    ' Reset checkboxes And toggle buttons To default state
    Me.VendorChx.value = False
    VendorChx_AfterUpdate
    Me.IsActive.value = True ' Default To "Active"

    ' Set focus To the first input field
    Me.PartyName.SetFocus

    On Error Goto 0 ' Reset error handling
End Sub

Private Sub AgreementExpiry_GotFocus()
    InputDateField AgreementExpiry, "Select a date To use this on your form"
End Sub
Private Sub Form_Load()
    Me.VendorChx.value = 0
    VendorChx_AfterUpdate
End Sub

Private Sub VendorChx_AfterUpdate()
    If Me.VendorChx.value = 0 Then
        Me.VendorPrimaryEmail.Visible = False
        Me.VendorRating.Visible = False
    Else

        Me.VendorRating.Visible = True
        Me.VendorPrimaryEmail.Visible = True
    End If

    If Me.VendorChx.value = -1 Then
        MsgBox "Please add the Vendor Primary email, it might be different"
        Me.VendorPrimaryEmail.Visible = True
        Me.VendorRating.Visible = True
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
    DoCmd.Close acForm, "AddClientF_V13"

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





' Vendor/Client Management System Documentation
' Overview
' This Access VBA module provides a complete solution For managing vendor And client information With robust transaction handling, validation, And logging capabilities. The system handles:

' Party creation (vendors/clients)

' Location management

' Contact information

' Role assignments

' Detailed vendor/client-specific data

' Core Components
' 1. Logging System
' LogError Procedure

' Purpose: Centralized error And activity logging

' Parameters:

' ErrorMessage: Description of the error/event

' ErrorSource: Source of the error

' ErrorNumber: Error number (default 0)

' ModuleName: Module where error occurred

' ProcedureName: Procedure where error occurred

' AdditionalInfo: Extra context information

' Log Format:

' Timestamp

' Module/Procedure details

' Error details

' Additional context

' Output: Writes To "TransactionAudit.log" in application directory

' 2. Validation Framework
' ValidateRequiredFields Function

' Validates all mandatory form fields

' Returns: Name of first missing/invalid field Or empty string If valid

' Uses helper functions:

' GetRequiredFieldsCollection: Builds collection of required fields

' IsFieldMissing: Checks If a field is empty/invalid

' HandleMissingField: Displays error And focuses missing field

' IsValidEmail: Validates email format With regex

' 3. Transaction Management
' AddPartyBtn_Click Procedure

' Main entry point For adding New parties

' Implements complete transaction workflow:

' Validates required fields

' Begins transaction

' Executes sequential operations:

' AddParty

' AddLocation

' AddPartyLocation

' AddPartyContact

' AddPartyDetails (vendor/client specific)

' AddPartyRole

' Commits on success Or rolls back On Error

' ExecuteInTransaction Function

' Executes specified operation within transaction context

' Parameters:

' operation: Name of operation To execute

' params: Array of parameters For the operation

' Returns: Operation result Or -1 on failure

' Logs operation start/end And parameters

' 4. Database Operations
' Party Management
' AddParty Function

' Creates New party record

' Generates unique CorpID And PartyCode

' Validates For duplicate party names

' Returns: New CorpID Or -1 on failure

' PartyNameExists Function

' Checks If party name already exists

' Returns: Boolean indicating existence

' Location Management
' AddLocation Function

' Creates location record

' Sets first location As primary

' Returns: New location ID Or -1 on failure

' AddPartyLocation Function

' Links party To location

' Validates For duplicates

' Returns: Boolean success status

' Contact Management
' AddPartyContact Function

' Adds contact information For party

' Returns: Boolean success status

' Vendor/Client Specifics
' AddPartyDetails Function

' Adds vendor Or client specific details

' Parameters:

' PartyID: ID of party

' TableName: "VendorsT_V13" Or "ClientsT_V13"

' isVendor: Boolean indicating vendor status

' Handles different fields For vendors vs clients

' Returns: Boolean success status

' Role Management
' AddPartyRole Function

' Assigns roles To party

' Validates For duplicate assignments

' Returns: Boolean success status

' 5. Utility Functions
' Code Generation
' GeneratePartyCode Function

' Generates sequential CorpID from PartyCodeT_V13

' Returns: New CorpID Or -1 on failure

' GenerateDisplayPartyCode Function

' Creates formatted display code

' Format: "VC" + 6 digits + 2 digit year (vendors) Or "C" prefix (clients)

' Returns: Formatted code Or error code

' Validation Utilities
' IsDuplicate* Functions

' Check For duplicate relationships:

' PartyLocation

' VendorCountry

' PartyRole

' PartyServiceCategory

' All follow same pattern:

' Accept two ID parameters

' Return Boolean indicating existence

' Log results

' RecordExists Function

' Generic record existence check

' Parameters:

' TableName: Table To check

' fieldName: Field To check

' value: Value To match

' Returns: Boolean indicating existence

' 6. Form Management
' ResetForm Procedure

' Clears all form fields

' Resets To default values

' Sets focus To first field

' Event Handlers

' Form_Load: Initializes form

' VendorChx_AfterUpdate: Toggles vendor-specific fields

' Country_AfterUpdate: Updates city list based on country

' *_GotFocus: Various field focus handlers

' Button handlers For actions (Add, Reset, Close)

' Database Schema
' Key tables used:

' PartiesT_V13: Core party information

' LocationsT_V13: Physical locations

' PartyLocations_JT_V13: Party-location relationships

' PartyContactT_V13: Contact information

' VendorsT_V13: Vendor-specific data

' ClientsT_V13: Client-specific data

' RolesT_V13: Role definitions

' PartyRolesJT_V13: Party-role assignments

' PartyCodeT_V13: CorpID sequence tracker

' Error Handling
' Comprehensive error handling throughout

' All procedures include:

' Error handler section

' Cleanup section For resource management

' Consistent logging

' Transaction rollback on any error

' Usage Flow
' User enters data in form

' System validates all required fields

' Transaction begins

' System sequentially creates:

' Party record

' Location record

' Party-location relationship

' Contact information

' Vendor/client specific records

' Role assignments

' Transaction commits on success Or rolls back on failure

' User receives success/failure notification

' Form resets For Next entry

' Best Practices Implemented
' Transaction Safety:

' All-Or-nothing operations

' Proper rollback handling

' Data Validation:

' Field-level validation

' Business rule enforcement

' Duplicate prevention

' Error Handling:

' Consistent structure

' Detailed logging

' User-friendly messages

' Resource Management:

' Proper recordset cleanup

' Memory management

' Auditability:

' Comprehensive logging

' Timestamped operations

' Contextual information   