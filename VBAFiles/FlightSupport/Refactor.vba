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

'=======================================================
' Validation Functions
'=======================================================
Private Function ValidateRequiredFields() As String
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "ValidateRequiredFields"

        Dim requiredFields As Collection
        Set requiredFields = GetRequiredFieldsCollection()

        LogError "Number of required fields: " & requiredFields.Count, "INFO", 0, "Validation", PROC_NAME

        If requiredFields.Count = 0 Then
            LogError "Required fields collection is empty", "CRITICAL", 0, "Validation", PROC_NAME
            ValidateRequiredFields = "Unknown field (validation error)"
         Exit Function
        End If

        Dim field As Variant
        For Each field In requiredFields
            LogError "Validating field: " & field(1), "INFO", 0, "Validation", PROC_NAME
            If IsFieldMissing(field) Then
                HandleMissingField field
                ValidateRequiredFields = field(1)
             Exit Function
            End If
        Next field

        ValidateRequiredFields = ""
     Exit Function

 ErrorHandler:
        LogError "Error in ValidateRequiredFields: " & Err.Description & " | Line: " & Erl, "CRITICAL", Err.Number, "Validation", PROC_NAME
        ValidateRequiredFields = "Unknown field (validation error)"
End Function

Private Function GetRequiredFieldsCollection() As Collection
    On Error Goto ErrorHandler
        Dim requiredFields As Collection
        Set requiredFields = New Collection

        requiredFields.Add Array(Me.BasePrice, "Base Price")
        requiredFields.Add Array(Me.EstimatedTotaltxt, "Estimated Total")
        requiredFields.Add Array(Me.Airport, "Airport")
        requiredFields.Add Array(Me.Vendor, "Vendor")
        requiredFields.Add Array(Me.Destination, "Destination")
        requiredFields.Add Array(Me.FlightType, "Flight Type")
        requiredFields.Add Array(Me.RampName, "Ramp Name")
        requiredFields.Add Array(Me.Currency, "Currency")
        requiredFields.Add Array(Me.Unit, "Unit")
        requiredFields.Add Array(Me.STC, "Subject To Change Checkbox")
        requiredFields.Add Array(Me.ChangeCyclebx, "Change Cycle")
        requiredFields.Add Array(Me.FuelType, "Fuel Type")
        requiredFields.Add Array(Me.Validity, "Validity")
        requiredFields.Add Array(Me.Notes, "Notes")
        requiredFields.Add Array(Me.FuelingMethodTxt, "Fueling Method")
        requiredFields.Add Array(Me.RangeTxt, "Range")

        If Me.OperatorCbx.Value Then
            requiredFields.Add Array(Me.Nationality, "Operator Nationality")
        End If

        Set GetRequiredFieldsCollection = requiredFields
     Exit Function

 ErrorHandler:
        LogError "Error in GetRequiredFieldsCollection: " & Err.Description, "CRITICAL", Err.Number, "Validation", "GetRequiredFieldsCollection"
        Set GetRequiredFieldsCollection = New Collection
End Function

Private Function IsFieldMissing(field As Variant) As Boolean
    On Error Goto ErrorHandler
        Dim fieldControl As Control
        Set fieldControl = field(0)

        If TypeOf fieldControl Is ListBox Then
            IsFieldMissing = (fieldControl.ItemsSelected.Count = 0)
        Elseif TypeOf fieldControl Is CheckBox Then
            IsFieldMissing = IsNull(fieldControl.Value)
        Else
            IsFieldMissing = (IsNull(fieldControl.Value) Or fieldControl.Value = "")
        End If

     Exit Function

 ErrorHandler:
        LogError "Error in IsFieldMissing: " & Err.Description, "CRITICAL", Err.Number, "Validation", "IsFieldMissing"
        IsFieldMissing = True
End Function

Private Sub HandleMissingField(field As Variant)
    On Error Goto ErrorHandler
        Dim fieldControl As Control
        Dim fieldName As String

        Set fieldControl = field(0)
        fieldName = field(1)

        If TypeOf fieldControl Is ListBox Then
            MsgBox "You must Select at least one item in '" & fieldName & "'.", vbExclamation, "Missing Required Selection"
        Elseif TypeOf fieldControl Is CheckBox Then
            MsgBox "The checkbox '" & fieldName & "' must be checked.", vbExclamation, "Missing Required Checkbox"
        Else
            MsgBox "The field '" & fieldName & "' is required. Please fill it in.", vbExclamation, "Missing Required Field"
        End If

        fieldControl.SetFocus
     Exit Sub

 ErrorHandler:
        LogError "Error in HandleMissingField: " & Err.Description, "CRITICAL", Err.Number, "Validation", "HandleMissingField"
End Sub

'=======================================================
' Add Fuel Price Details
'=======================================================
Private Sub AddFuelPriceBtn_Click()
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddFuelPriceBtn_Click"
        Dim ws As DAO.Workspace
        Dim success As Boolean
        Dim missingField As String

        LogError "Process initialization", "INFO", 0, "Form", PROC_NAME

        missingField = ValidateRequiredFields()
        If missingField <> "" Then
            LogError "Validation failed. Missing Or invalid field: " & missingField, "VALIDATION", 0, "Form", PROC_NAME
         Exit Sub
        End If

        Set ws = DBEngine.Workspaces(0)
        ws.BeginTrans
        On Error Goto TransactionError

            success = ExecuteInTransaction("Add_DropDowns")
            If success = -1 Then Goto Rollback

                success = ExecuteInTransaction("Add_FuelPriceDetails")
                If success = -1 Then Goto Rollback

                    ws.CommitTrans
                    LogError "Transaction completed successfully", "SUCCESS", 0, "Form", PROC_NAME
                    MsgBox "Operation completed successfully!", vbInformation
                    ResetForm
                 Exit Sub

 TransactionError:
                    LogError "Transaction failure: " & Err.Description, "CRITICAL", Err.Number, "Form", PROC_NAME
 Rollback:
                    ws.Rollback
                    MsgBox "Operation failed. Check audit log.", vbExclamation
                 Exit Sub

 ErrorHandler:
                    LogError "Unexpected error: " & Err.Description, "CRITICAL", Err.Number, "Form", PROC_NAME
                    Resume Rollback
End Sub

Private Function ExecuteInTransaction(Byval operation As String, ParamArray params() As Variant) As Variant
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "ExecuteInTransaction"

        Select Case operation
         Case "Add_DropDowns"
            ExecuteInTransaction = Add_DropDowns()
         Case "Add_FuelPriceDetails"
            ExecuteInTransaction = Add_FuelPriceDetails()
         Case Else
            LogError "Invalid operation: " & operation, "ERROR", 9001, "Transaction", PROC_NAME
            ExecuteInTransaction = -1
        End Select

     Exit Function

 ErrorHandler:
        LogError "Operation failed: " & Err.Description, "CRITICAL", Err.Number, "Transaction", PROC_NAME
        ExecuteInTransaction = -1
End Function

    Private Sub Add_DropDowns()
    On Error Goto ErrorHandler
        Dim rs As DAO.Recordset
        Dim selectedAirport As Integer
        Dim selectedFlightType As Integer
        Dim selectedVendor As Integer
        Dim selectedDestination As Integer
        Dim selectedRampName As Integer
        Dim selectedFuelingMethod As Integer

        selectedAirport = Me.Airport.Column(0)
        selectedFlightType = Me.FlightType.Column(0)
        selectedVendor = Me.Vendor.Column(0)
        selectedDestination = Me.Destination.Column(0)
        selectedRampName = Me.RampName.Column(0)
        selectedFuelingMethod = Me.FuelingMethodTxt.Column(0)

        Set rs = CurrentDb.OpenRecordset("FuelSupport_FuelLocations_JT_V13", dbOpenDynaset, dbAppendOnly)
        ' Check If the record already exists
        rs.FindFirst "Airport_ID = " & selectedAirport & " And FightType_ID = " & selectedFlightType & " And Vendor_ID = " & selectedVendor & " And Dest_ID = " & selectedDestination & ""
        If Not rs.NoMatch Then
            MsgBox "Record already exists."
            Goto CleanExit
            End If
            ' Add New record
            With rs
                .AddNew
                !Airport_ID = selectedAirport
                !AirportCode = Me.Airport.Column(1)

                !FightType_ID = selectedFlightType
                !FightTypeName = Me.FlightType.Column(1)

                !Vendor_ID = selectedVendor
                !vendorName = Me.Vendor.Column(1)

                !Ramp_ID = selectedRampName
                !RampName = Me.RampName.Column(1)

                !FuelType_ID = Me.FuelType.Column(0)
                !FuelTypeName = Me.FuelType.Column(1)

                !Dest_ID = Me.Destination.Column(0)
                !DestName = Me.Destination.Column(1)

                !FuelMethod_ID = Me.FuelingMethodTxt.Column(0)
                !FuelMethod = Me.FuelingMethodTxt.Column(1)

                !FBO_ID = Me.FBO.Column(0)
                !FBOName = Me.FBO.Column(1)

                !IPA_ID = Me.IPA.Column(0)
                !IPAName = Me.IPA.Column(1)

                !Unit_ID = Me.Unit.Column(0)
                !Unit = Me.Unit.Column(1)
                !Currency_ID = Me.Currency.Column(0)
                !Currency = Me.Currency.Column(1)

                .Update
            End With

            ' Reset the form

            MsgBox "Drop downs added successfully."
            Me.Undo

 CleanExit:
            On Error Resume Next
            If Not rs Is Nothing Then
                rs.Close
                Set rs = Nothing
            End If
            Set db = Nothing
         Exit Sub

 ErrorHandler:
            MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error"
            Resume CleanExit
End Sub
Private Sub Add_FuelPriceDetails()
    On Error Goto ErrorHandler
        Dim rs As DAO.Recordset
        Dim selectedAirport As Integer
        Dim selectedFlightType As Integer
        Dim selectedVendor As Integer
        Dim selectedDestination As Integer
        Dim selectedRampName As Integer
        Dim selectedFuelingMethod As Integer

        selectedAirport = Me.Airport.Column(0)
        selectedFlightType = Me.FlightType.Column(0)
        selectedVendor = Me.Vendor.Column(0)
        selectedDestination = Me.Destination.Column(0)
        selectedRampName = Me.RampName.Column(0)
        selectedFuelingMethod = Me.FuelingMethodTxt.Column(0)

        Set rs = CurrentDb.OpenRecordset("FuelSupport_FuelPricesT_V13", dbOpenDynaset, dbAppendOnly)
        ' Check If the record already exists
        rs.FindFirst "Airport_ID = " & selectedAirport & " And FightType_ID = " & selectedFlightType & " And Vendor_ID = " & selectedVendor & " And Dest_ID = " & selectedDestination & ""
        If Not rs.NoMatch Then
            MsgBox "Record already exists."
            Goto CleanExit
            End If

            ' Add New record
            With rs
                .AddNew
                !Airport_ID = selectedAirport
                !FightType_ID = selectedFlightType
                !Vendor_ID = selectedVendor
                !Ramp_ID = selectedRampName
                !FuelType_ID = Me.FuelType.Column(0)
                !Dest_ID = Me.Destination.Column(0)
                !FuelMethod_ID = Me.FuelingMethodTxt.Column(0)
                !FBO_ID = Me.FBO.Column(0)
                !IPA_ID = Me.IPA.Column(0)
                If Not IsNull(Me.Nationality.value) Or Not IsNull(Me.Operator.value) Then
                    !Operator_ID = Me.Operator.Column(0)
                End If

                !BasePrice = Me.BasePrice.value
                !EstimatedTotal = Me.EstimatedTotaltxt.value

                !Notes = Me.Notes.value
                !Taxes = Me.Taxes.value
                !InternalNotes = Me.InternalNotes.value

                !AddedBy = Me.Usertxt.value
                !Validity = Me.Validity.value
                !ChangeCycle = Me.ChangeCyclebx.value
                !SubjectToChange = Me.STC.value
                !Range = Me.RangeTxt.value

                .Update
            End With

            ' Reset the form

            MsgBox "fuel price details added successfully."
            Me.Undo

 CleanExit:
            On Error Resume Next
            If Not rs Is Nothing Then
                rs.Close
                Set rs = Nothing
            End If
            Set db = Nothing
         Exit Sub

 ErrorHandler:
            MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error"
            Resume CleanExit
End Sub

