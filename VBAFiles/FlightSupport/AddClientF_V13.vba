Option Compare Database
Option Explicit
Private Sub City_GotFocus()
    If IsNull(Me.Country.value) Then
        MsgBox "Select a country first"
        Me.Country.SetFocus
    End If
End Sub

Private Sub CloseBtn_Click()
    'ResetFormFields
    Me.Undo
    DoCmd.Close acForm, "AddClientF_V13"
End Sub
Private Sub Country_AfterUpdate()
    Dim rst As DAO.Recordset
    Dim Nationality As DAO.Recordset
    Dim selectedCountry As Variant
    Dim nationalityValue As String
    Dim sql As String

    On Error Goto ErrorHandler

        ' Get the selected country code
        selectedCountry = Me.Country.Column(0)

        ' Sanitize And build the SQL query For nationality
        sql = SanitizeSQL("Select nationality FROM CountriesT_V13 WHERE num_code = [param1];", selectedCountry)
        Set Nationality = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

        ' Check If the recordset is Not empty
        If Not Nationality.EOF Then
            nationalityValue = Nationality!Nationality
        Else
            nationalityValue = "" ' Or some default value
        End If

        ' Set the nationality value in the form
        'Me.Nationality.Value = nationalityValue

        ' Close the nationality recordset
        Nationality.Close
        Set Nationality = Nothing

        ' Sanitize And build the SQL query For cities
        sql = SanitizeSQL("Select City FROM CitiesT_V13 WHERE CountryCode = [param1];", selectedCountry)
        Me.City.RowSource = sql
        Me.City.Requery

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & Err.Description
        If Not Nationality Is Nothing Then
            Nationality.Close
            Set Nationality = Nothing
        End If
End Sub
' Reusable Function To sanitize SQL queries With parameters
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
Private Sub DateSigned_Click()
    InputDateField DateSigned, "Select a date To use this on your form"
End Sub
Private Sub AOCExpiry_Click()
    InputDateField AOCExpiry, "Select a date To use this on your form"
End Sub
Private Function ExecuteInTransaction(operation As String, Optional param1 As Long, Optional param2 As Long) As Variant
    On Error Goto ErrorHandler

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
            AddPartyContact param1, param2
            ExecuteInTransaction = True
            LogError "AddPartyContact completed.", "ExecuteInTransaction"

         Case "AddClient"
            AddClient param1
            ExecuteInTransaction = True
            LogError "AddClient completed.", "ExecuteInTransaction"

         Case "AddVendor"
            AddVendor param1
            ExecuteInTransaction = True
            LogError "AddVendor completed.", "ExecuteInTransaction"

         Case "AddPartyLocation"
            AddPartyLocation param1, param2
            ExecuteInTransaction = True
            LogError "AddPartyLocation completed.", "ExecuteInTransaction"

         Case "AddPartyRole"
            AddPartyRole param1, param2
            ExecuteInTransaction = True
            LogError "AddPartyRole completed.", "ExecuteInTransaction"

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
'=======================================================
' Helper Function: GeneratePartyCode
'=======================================================
Private Function GeneratePartyCode() As Long
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Dim newCorpID As Long

        ' Step 1: Retrieve the last CorpID from the PartyCodeT_V13 table
        Set rst = CurrentDb.OpenRecordset("PartyCodeT_V13", dbOpenDynaset)

        If Not rst.EOF And Not IsNull(rst!CorpID) Then
            newCorpID = rst!CorpID + 1 ' Increment the last CorpID
        Else
            newCorpID = 1 ' Initialize To 1 If no records exist
        End If

        ' Step 2: Update the PartyCodeT_V13 table With the New CorpID
        If rst.recordCount > 0 Then
            rst.Edit ' Update the existing record
        Else
            rst.AddNew ' Add a New record If the table is empty
        End If
        rst!CorpID = newCorpID
        rst.Update

        rst.Close
        Set rst = Nothing

        ' Return the New CorpID
        GeneratePartyCode = newCorpID
     Exit Function

 ErrorHandler:
        LogError "Error in GeneratePartyCode: " & Err.Description, "GeneratePartyCode"
        MsgBox "Error in GeneratePartyCode: " & Err.Description, vbCritical
        GeneratePartyCode = -1
End Function
'=======================================================
' Helper Function: AddParty
'=======================================================
Private Function AddParty() As Long
    On Error Goto ErrorHandler

        ' Validate required fields
        If IsNull(Me.PartyName.value) Or IsNull(Me.FSANumber.value) Then
            MsgBox "Party Name And FSANumber are required.", vbExclamation
            AddParty = -1
         Exit Function
        End If

        Dim rst As DAO.Recordset
        Dim CorpID As Long
        Dim PartyCode As String

        ' Step 1: Generate the Next CorpID
        CorpID = GeneratePartyCode()
        If CorpID = -1 Then Exit Function ' Exit If there was an error

            ' Step 2: Generate the PartyCode dynamically
            If Me.VendorChx.value = True Then
                PartyCode = "CV" & Format(CorpID, "0000") & Format(date, "yy") ' Format As CVXXXXYY
            Else
                PartyCode = "C" & Format(CorpID, "0000") & Format(date, "yy") ' Format As CXXXXYY
            End If

            ' Step 3: Add the New party To the PartiesT_V13 table
            Set rst = CurrentDb.OpenRecordset("PartiesT_V13", dbOpenDynaset)

            With rst
                .AddNew
                !CorpID = CorpID
                !PartyCode = PartyCode
                !PartyName = Me.PartyName.value
                !FSANumber = Me.FSANumber.value
                !AgreementStatus = Me.AgreementStatus.value
                !AgreementExpiry = Me.AgreementExpiry.value ' One year date from today
                !DateSigned = Me.DateSigned.value
                !VAT_NO = Me.VAT_NO.value
                ' !DocStorage = Me.DocStorage.Value ' Uncomment If needed

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
'=======================================================
'This Function generates the PartyCode dynamically For display purposes
'=======================================================
Private Function GenerateDisplayPartyCode(CorpID As Long) As String
    Dim prefix As String

    ' Determine the prefix based on VendorChx
    If Me.VendorChx.value = True Then
        prefix = "CV" ' Use "CV" For clients that are also vendors
    Else
        prefix = "C" ' Use "C" For clients
    End If

    ' Generate the PartyCode in the format [Prefix][XXXX][YY]
    GenerateDisplayPartyCode = prefix & Format(CorpID, "0000") & Format(date, "yy")
End Function

'=======================================================
'Displays the dynamically generated PartyCode on the form.
'=======================================================
Private Sub DisplayPartyCode(CorpID As Long)
    Me.PartyCode.value = GenerateDisplayPartyCode(CorpID)
End Sub

'=======================================================
' Main Function: AddClientBtn_Click
'=======================================================
Private Sub AddPartyBtn_Click()
    On Error Goto ErrorHandler

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
        If PartyID = -1 Then Goto Rollback
            LogError "AddParty completed. PartyID: " & PartyID, "AddPartyBtn_Click"

            ' Add Location
            locationID = AddLocation(PartyID)
            If locationID = -1 Then Goto Rollback
                LogError "AddLocation completed. LocationID: " & locationID, "AddPartyBtn_Click"

                ' Add Party Contact
                AddPartyContact PartyID, locationID
                LogError "AddPartyContact completed.", "AddPartyBtn_Click"

                ' Add Client
                AddClient PartyID
                LogError "AddClient completed.", "AddPartyBtn_Click"

                ' Add Roles (from ListBox)
                Dim RoleID As Long
                Dim i As Variant
                For Each i In Me.RolesListBox.ItemsSelected
                    RoleID = Me.RolesListBox.Column(0, i) ' Assuming the Role ID is in the first column
                    AddPartyRole PartyID, RoleID
                    LogError "AddPartyRole completed For RoleID: " & RoleID, "AddPartyBtn_Click"
                Next i

                ' Add Vendor (If applicable)
                If Me.VendorChx.value = True Then
                    ' Add Vendor
                    AddVendor PartyID
                    LogError "AddVendor completed.", "AddPartyBtn_Click"

                    ' Add Party Location (Vendor)
                    AddPartyLocation PartyID, locationID
                    LogError "AddPartyLocation completed.", "AddPartyBtn_Click"
                End If

                ' Commit the transaction If everything succeeds
                ws.CommitTrans
                LogError "Transaction committed successfully.", "AddPartyBtn_Click"
                MsgBox "Client added successfully!", vbInformation
             Exit Sub

 Rollback:
                ' Rollback the transaction On Error
                ws.Rollback
                LogError "Transaction rolled back due To an error.", "AddPartyBtn_Click"
                MsgBox "Transaction rolled back due To an error.", vbExclamation
             Exit Sub

 ErrorHandler:
                ' Log the error And rollback the transaction
                LogError "Error in AddPartyBtn_Click: " & Err.Description, "AddPartyBtn_Click"
                MsgBox "Error in AddPartyBtn_Click: " & Err.Description, vbCritical
                If Not ws Is Nothing Then
                    ws.Rollback ' Ensure the transaction is rolled back
                    LogError "Transaction rolled back due To an error.", "AddPartyBtn_Click"
                End If
End Sub
'=======================================================
' Helper Function: AddPartyRole
'=======================================================
Private Function AddPartyRole(PartyID As Long, RoleID As Long) As Boolean
    On Error Goto ErrorHandler

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

'=======================================================
' Helper Function: AddLocation
'=======================================================
Private Function AddLocation(PartyID As Long) As Long
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("LocationsT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !CorpID = PartyID ' Use the passed PartyID
            !Country = Me.Country.value
            !City = Me.City.value
            !Address = Me.Address.value
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

'=======================================================
' Helper Function: AddPartyContact
'=======================================================
Private Sub AddPartyContact(PartyID As Long, locationID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("PartyContactT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !CorpID = PartyID
            !locationID = locationID
            !PrimaryEmail = Me.PrimaryEmail.value
            !SecondaryEmail = Me.SecondaryEmail.value
            !Phone = Me.Phone.value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddPartyContact: " & Err.Description, "AddPartyContact"
        MsgBox "Error in AddPartyContact: " & Err.Description, vbCritical
End Sub

'=======================================================
' Helper Function: AddClient
'=======================================================
Private Sub AddClient(PartyID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("ClientsT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !CorpID = PartyID
            !CreditLimit = Me.CreditLimit.value
            !Deposit = Me.Deposit.value
            !Currency = Me.Currency.value
            !ClientRating = Me.ClientRating.value
            !ClientDescription = Me.ClientDescription.value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddClient: " & Err.Description, "AddClient"
        MsgBox "Error in AddClient: " & Err.Description, vbCritical
End Sub

'=======================================================
' Helper Function: AddVendor
'=======================================================
Private Sub AddVendor(PartyID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("VendorsT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !CorpID = PartyID
            !CreditLimit = Me.CreditLimit.value
            !Deposit = Me.Deposit.value
            !Currency = Me.Currency.value
            !VendorRating = Me.VendorRating.value
            !VendorDescription = Me.ClientDescription.value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddVendor: " & Err.Description, "AddVendor"
        MsgBox "Error in AddVendor: " & Err.Description, vbCritical
End Sub

'=======================================================
' Helper Function: AddPartyLocation
'=======================================================
Private Sub AddPartyLocation(PartyID As Long, locationID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("PartyLocations_JT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !CorpID = PartyID
            !locationID = locationID
            !isPrimary = True
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddPartyLocation: " & Err.Description, "AddPartyLocation"
        MsgBox "Error in AddPartyLocation: " & Err.Description, vbCritical
End Sub

'=======================================================
' Helper Function: logging errors
'=======================================================
Private Sub LogError(ErrorMessage As String, ErrorSource As String)
    On Error Goto ErrorHandler ' Temporarily remove "On Error Resume Next"

        Dim filePath As String
        Dim fileNumber As Integer
        Dim logMessage As String

        ' Define the file path
        filePath = "E:\ent\DebugLog_AddClientF_V13.txt"
        'MsgBox "Log file path: " & filePath, vbInformation ' Debug message

        ' Open the file For appending (creates the file If it doesn't exist)
        fileNumber = FreeFile
        ' MsgBox "Opening file For appending...", vbInformation ' Debug message
        Open filePath For Append As #fileNumber
        ' MsgBox "File opened successfully.", vbInformation ' Debug message

        ' Prepare the log message
        logMessage = "Error Date: " & Now() & vbCrLf & _
        "Error Source: " & ErrorSource & vbCrLf & _
        "Error Message: " & ErrorMessage & vbCrLf & _
        "----------------------------------------" & vbCrLf

        ' Write the log message To the file
        'MsgBox "Writing log message: " & logMessage, vbInformation ' Debug message
        Print #fileNumber, logMessage
        ' MsgBox "Log message written successfully.", vbInformation ' Debug message

        ' Close the file
        Close #fileNumber
        ' MsgBox "File closed successfully.", vbInformation ' Debug message

     Exit Sub

 ErrorHandler:
        MsgBox "Error in LogError: " & Err.Description, vbCritical ' Debug message
End Sub
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
    Me.ClientDescription.value = Null
    Me.VendorRating.value = 3 ' 3 is the default value


    ' Reset checkboxes And toggle buttons To default state
    Me.VendorChx.value = False
    Me.IsActive.value = True ' Default To "Active"

    ' Set focus To the first input field
    Me.PartyName.SetFocus

    On Error Goto 0 ' Reset error handling
End Sub
Private Sub ResetBtn_Click()
    ' Ask the user For confirmation before resetting the form
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want To reset the form? All unsaved data will be lost.", vbYesNo + vbQuestion, "Confirm Reset")

    ' Check the user's response
    If response = vbYes Then
        ResetForm ' Reset the form If the user clicks Yes
        LogError "Form reset by user.", "ResetBtn_Click" ' Optional logging
    Else
        ' Do nothing If the user clicks No
     Exit Sub
    End If
End Sub
Private Sub Command192_Click()
    DoCmd.OpenForm "UpdateOperatorF", WindowMode:=acWindowNormal
    CloseBtn_Click
End Sub
Private Sub Form_Load()
    Dim fw As New clFormWindow
    DoCmd.GoToRecord , , acNewRec
    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
    Me.VendorRating.Visible = False

End Sub
Private Sub VendorChx_Click()
    If Me.VendorChx.value = -1 Then
        Me.VendorRating.Visible = True
    Else
        Me.VendorRating.Visible = False
    End If
End Sub