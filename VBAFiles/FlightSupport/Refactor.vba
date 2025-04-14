Private Sub AddPartyBtn_Click()
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyBtn_Click"
        Dim ws As DAO.Workspace
        Dim PartyID As Long, locationID As Long
        Dim success As Boolean, CountryCode As Long
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
                            success = ExecuteInTransaction("AddPartyContact", PartyID,locationID, RoleID)
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
                                        success = ExecuteInTransaction("AddPartyRole", PartyID, RoleID)
                                        If Not success Then Goto Rollback
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
            ExecuteInTransaction = AddPartyRole(CLng(params(0)), CLng(params(1)))
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

        ' Add record (simplified - no IsActive field)
        LogError "Opening PartyRolesJT_V13", "INFO", 0, "Database", PROC_NAME
        Set rst = CurrentDb.OpenRecordset("PartyRolesJT_V13", dbOpenDynaset)
        With rst
            LogError "Adding New record", "INFO", 0, "Database", PROC_NAME
            .AddNew
            !corpID = PartyID
            !RoleID = RoleID
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
Private Function AddPartyContact( _
    PartyID As Long, _
    LocationID As Long, _
    RoleID As Long _
    ) As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddPartyContact"
        Dim rst As DAO.Recordset
        Dim isVendor As Boolean
        Dim isClient As Boolean

        ' === VALIDATION ===
        If PartyID <= 0 Or LocationID <= 0 Or RoleID <= 0 Then
            LogError "Invalid IDs provided", "VALIDATION", 4001, "Database", PROC_NAME
         Exit Function
        End If

        ' === ROLE DETECTION ===
        isVendor = (RoleID = GetVendorRoleID())
        isClient = (RoleID = GetClientRoleID())

        ' Check For existing contact For this role
        If RecordExists("PartyContactT_V13", "corpID = " & PartyID & " And RoleID = " & RoleID) Then
            LogError "Duplicate contact For RoleID " & RoleID, "VALIDATION", 4003, "Database", PROC_NAME
         Exit Function
        End If

        ' === DATABASE OPERATION ===
        Set rst = CurrentDb.OpenRecordset("PartyContactT_V13", dbOpenDynaset)

        With rst

            .AddNew
            !corpID = PartyID
            !locationID = LocationID
            !RoleID = RoleID ' Store 

            ' --- Set contact info based on role ---
            Select Case True
             Case isVendor:
                !PrimaryEmail = Nz(Me.VendorPrimaryEmail.Value, "")
                !SecondaryEmail = Nz(Me.SecondaryEmail.Value, "")
                !Phone = Nz(Me.Phone.Value, "")
                !ContactType = "VENDOR"

             Case isClient:
                !PrimaryEmail = Nz(Me.ClientPrimaryEmail.Value, "")
                !SecondaryEmail = Nz(Me.SecondaryEmail.Value, "")
                !Phone = Nz(Me.Phone.Value, "")
                !ContactType = "CLIENT"

            End Select

            .Update
        End With

        LogError "Contact added For " & IIf(RoleID = 0, "default", "RoleID " & RoleID), _
        "SUCCESS", 0, "Database", PROC_NAME
        AddPartyContact = True

 Cleanup:
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
     Exit Function

 ErrorHandler:
        LogError "Failed To add contact For RoleID " & RoleID & ": " & Err.Description, _
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