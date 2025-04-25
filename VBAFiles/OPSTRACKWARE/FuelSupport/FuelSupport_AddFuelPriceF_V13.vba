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
    Optional Byval AdditionalInfo As String = "", _
    Optional Byval LineNumber As Long = 0 _
    )
    On Error Goto LogError_Handler
        Const LOG_FILE_NAME As String = "\TransactionAudit.log"

        Dim filePath As String
        Dim fileNumber As Integer
        Dim logMessage As String

        ' Create ISO 8601 timestamp
        logMessage = String(50, "=") & vbCrLf & _
        "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "Module: " & ModuleName & vbCrLf & _
        "Procedure: " & ProcedureName & vbCrLf & _
        "Error #: " & ErrorNumber & vbCrLf & _
        "Source: " & ErrorSource & vbCrLf & _
        "Message: " & ErrorMessage & vbCrLf & _
        "Line #: " & LineNumber & vbCrLf & _
        "Additional Info: " & AdditionalInfo & vbCrLf & _
        String(50, "=") & vbCrLf & vbCrLf

        ' Secure file handling
        filePath = CurrentProject.Path & LOG_FILE_NAME
        fileNumber = FreeFile
        Open filePath For Append As #fileNumber
        Print #fileNumber, logMessage
        Close #fileNumber

     Exit Sub

 LogError_Handler:
        MsgBox "Logging Failed: " & Err.Description, vbCritical
End Sub

'=======================================================
' Core Validation Functions
'=======================================================
Private Function ValidateAndFormatRange() As String
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "ValidateAndFormatRange"

        Dim rawRange As String
        Dim parts() As String
        Dim startVal As Long
        Dim endVal As Long

        rawRange = Nz(Me.RangeTxt.value, "")
        rawRange = Replace(rawRange, " ", "")

        ' Validate format
        If Not rawRange Like "#*-#*" Then
            LogError "Invalid range format: " & rawRange, "ERROR", , "Validation", PROC_NAME
            MsgBox "Range format: start-end (e.g., 1-19999)", vbExclamation
            ValidateAndFormatRange = "INVALID"
         Exit Function
        End If

        ' Split And validate
        parts = Split(rawRange, "-")
        If UBound(parts) <> 1 Then
            LogError "Multiple hyphens: " & rawRange, "ERROR", , "Validation", PROC_NAME
            MsgBox "Use single hyphen separator", vbExclamation
            ValidateAndFormatRange = "INVALID"
         Exit Function
        End If

        ' Numeric validation
        If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
            LogError "Non-numeric values: " & rawRange, "ERROR", , "Validation", PROC_NAME
            MsgBox "Numbers only in range", vbExclamation
            ValidateAndFormatRange = "INVALID"
         Exit Function
        End If

        startVal = CLng(parts(0))
        endVal = CLng(parts(1))

        If startVal >= endVal Then
            LogError "Invalid range order: " & rawRange, "ERROR", , "Validation", PROC_NAME
            MsgBox "Start must be < end", vbExclamation
            ValidateAndFormatRange = "INVALID"
         Exit Function
        End If

        ValidateAndFormatRange = startVal & "-" & endVal
     Exit Function

 ErrorHandler:
        LogError "Range processing failed: " & Err.Description, "CRITICAL", Err.Number, "Validation", PROC_NAME
        ValidateAndFormatRange = "ERROR"
End Function

Private Function GetAirportCode(ctrl As Control) As String
    On Error Goto ErrorHandler
        Dim rawCode As String
        rawCode = Replace(Nz(ctrl.Column(1), " ", ""))

        ' Format validation And truncation
        If Not rawCode Like "[A-Z][A-Z][A-Z][A-Z][/]*" Then
            LogError "Airport code format: " & rawCode, "WARNING", , "Validation", "GetAirportCode"
        End If

        GetAirportCode = UCase(Left(rawCode, 10))
     Exit Function

 ErrorHandler:
        LogError "Airport code error", "CRITICAL", Err.Number, "Validation", "GetAirportCode"
        GetAirportCode = "INVALID"
End Function

'=======================================================
' Field Validation System
'=======================================================
Private Function ValidateRequiredFields() As String
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "ValidateRequiredFields"

        ' Airport validation
        If Not ValidateAirport() Then
            ValidateRequiredFields = "Airport"
         Exit Function
        End If

        ' Required fields check
        Dim col As Collection
        Set col = GetRequiredFieldsCollection()

        If col.Count = 0 Then
            LogError "Empty validation Set", "CRITICAL", , "Validation", PROC_NAME
            ValidateRequiredFields = "Configuration error"
         Exit Function
        End If

        Dim field As Variant
        For Each field In col
            If IsFieldMissing(field) Then
                HandleMissingField field
                ValidateRequiredFields = field(1)
             Exit Function
            End If
            Next

            ' Range validation
            Dim rangeResult As String
            rangeResult = ValidateAndFormatRange()
            If rangeResult = "INVALID" Then
                ValidateRequiredFields = "Range"
            Elseif rangeResult = "ERROR" Then
                ValidateRequiredFields = "System error"
            Else
                ValidateRequiredFields = ""
            End If

         Exit Function

 ErrorHandler:
            LogError "Validation failed: " & Err.Description, "CRITICAL", Err.Number, "Validation", PROC_NAME
            ValidateRequiredFields = "Validation error"
End Function

Private Function ValidateAirport() As Boolean
    On Error Goto ErrorHandler
        Dim code As String
        code = GetAirportCode(Me.Airport)

        ValidateAirport = (code Like "[A-Z][A-Z][A-Z][A-Z]") Or _
        (code Like "[A-Z][A-Z][A-Z][A-Z]/[A-Z][A-Z][A-Z]")

        If Not ValidateAirport Then
            MsgBox "Invalid airport code format", vbExclamation
            LogError "Invalid airport code: " & code, "ERROR", , "Validation", "ValidateAirport"
        End If
     Exit Function

 ErrorHandler:
        LogError "Airport validation failed", "CRITICAL", Err.Number, "Validation", "ValidateAirport"
        ValidateAirport = False
End Function

'=======================================================
' Database Operations
'=======================================================
Private Function Add_DropDowns() As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "Add_DropDowns"
        Const TABLE_NAME As String = "FuelSupport_FuelLocations_JT_V13"

        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb

        LogError "Starting dropdown insert", "INFO", , "Database", PROC_NAME

        ' Validate critical controls
        If GetID(Me.Airport) = 0 Or GetID(Me.FlightType) = 0 Then
            LogError "Invalid control selection", "ERROR", , "Database", PROC_NAME
            Add_DropDowns = False
         Exit Function
        End If

        Set rs = db.OpenRecordset(TABLE_NAME, dbOpenDynaset, dbAppendOnly)

        With rs
            .AddNew

            ' --- Airport Section ---
            !Airport_ID = CLng(GetID(Me.Airport))
            !airportCode = Left(GetAirportCode(Me.Airport), 10) ' Match TEXT(10)

            ' --- Flight Type ---
            !FlightType_ID = CLng(GetID(Me.FlightType))
            !FlightTypeName = Left(GetColumn(Me.FlightType, 1), 50)

            ' --- Vendor Details ---
            !Vendor_ID = CLng(GetID(Me.Vendor))
            !vendorName = Left(GetColumn(Me.Vendor, 1), 50)

            ' --- Ramp Information ---
            !Ramp_ID = CLng(GetID(Me.RampName))
            !RampName = Left(GetColumn(Me.RampName, 1), 50)

            ' --- Fuel Details ---
            !FuelType_ID = CLng(GetID(Me.FuelType))
            !FuelTypeName = Left(GetColumn(Me.FuelType, 1), 50)

            ' --- Fueling Method ---
            !FuelMethod_ID = CLng(GetID(Me.FuelingMethodTxt))
            !FuelMethod = Left(GetColumn(Me.FuelingMethodTxt, 1), 50)

            ' --- Destination ---
            !Dest_ID = CLng(GetID(Me.Destination))
            !DestName = Left(GetColumn(Me.Destination, 1), 50)

            ' --- Optional Fields ---
            !FBO_ID = CLng(Nz(Me.FBOIDtxt.value, 0))
            !FBOName = Left(Nz(Me.FBO.value, ""), 50)
            !IPA_ID = CLng(Nz(Me.IPAIDtxt.value, 0))
            !IPAName = Left(Nz(Me.IPA.value, ""), 50)

            ' --- Unit/Currency ---
            !Unit_ID = CLng(GetID(Me.Unit))
            !Unit = Left(GetColumn(Me.Unit, 1), 10)
            !Currency_ID = CLng(GetID(Me.Currency))
            !Currency = Left(GetColumn(Me.Currency, 1), 3)

            .Update
        End With

        Add_DropDowns = True
        Goto CleanExit

 ErrorHandler:
            LogError "Insert failed: " & Err.Description & " | Field: " & rs.Fields(rs.Fields.Count - 1).Name, _
            "CRITICAL", Err.Number, "Database", PROC_NAME
            Add_DropDowns = False

 CleanExit:
            If Not rs Is Nothing Then rs.Close
                Set rs = Nothing
                Set db = Nothing
End Function

Private Function Add_FuelPriceDetails() As Boolean
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "Add_FuelPriceDetails"
        Const TABLE_NAME As String = "FuelSupport_FuelPricesT_V13"

        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb

        LogError "Starting price insert", "INFO", , "Database", PROC_NAME

        ' Validate range
        Dim formattedRange As String
        formattedRange = ValidateAndFormatRange()
        If Left(formattedRange, 7) = "INVALID" Then
            LogError "Invalid range", "CRITICAL", , "Database", PROC_NAME
            Add_FuelPriceDetails = False
         Exit Function
        End If

        Set rs = db.OpenRecordset(TABLE_NAME, dbOpenDynaset, dbAppendOnly)

        With rs
            .AddNew
            ' Numeric values
            !Airport_ID = CLng(GetID(Me.Airport))
            !FlightType_ID = CLng(GetID(Me.FlightType))
            !Vendor_ID = CLng(GetID(Me.Vendor))
            !Ramp_ID = CLng(GetID(Me.RampName))
            !FuelMethod_ID = CLng(GetID(Me.FuelingMethodTxt))
            !Dest_ID = CLng(GetID(Me.Destination))

            ' --- Optional References ---
            !FBO_ID = CLng(Nz(Me.FBOIDtxt.value, 0))
            !IPA_ID = CLng(Nz(Me.IPAIDtxt.value, 0))

            ' --- Pricing Details ---
            !BasePrice = CDbl(Nz(Me.BasePrice.value, 0))
            !EstimatedTotal = CDbl(Nz(Me.EstimatedTotaltxt.value, 0))
            !Taxes = CDbl(Nz(Me.Taxes.value, 0))

            ' --- Text Fields ---
            !Notes = Left(Nz(Me.Notes.value, ""), 255)
            !InternalNotes = Left(Nz(Me.InternalNotes.value, ""), 255)
            !AddedBy = Left(Nz(Me.Usertxt.value, ""), 50)

            ' --- Date/Status ---
            If IsDate(Me.Validity.value) Then
                !Validity = CDate(Me.Validity.value)
            Else
                !Validity = Null
            End If
            !ChangeCycle = Left(Nz(Me.ChangeCyclebx.value, ""), 50)
            !SubjectToChange = CBool(Nz(Me.STC.value, False))

            ' --- Range Handling ---
            !Range = Left(formattedRange, 20)

            .Update
        End With

        Add_FuelPriceDetails = True
        Goto CleanExit

 ErrorHandler:
            LogError "Insert failed: " & Err.Description & " | Field: " & rs.Fields(rs.Fields.Count - 1).Name, _
            "CRITICAL", Err.Number, "Database", PROC_NAME
            Add_FuelPriceDetails = False

 CleanExit:
            If Not rs Is Nothing Then rs.Close
                Set rs = Nothing
                Set db = Nothing
End Function

'=======================================================
' Supporting Functions
'=======================================================
Private Function GetID(ctrl As Control) As Long
    On Error Goto ErrorHandler
        If ctrl.ListIndex = -1 Then Exit Function
            GetID = CLng(ctrl.Column(0))
         Exit Function

 ErrorHandler:
            LogError "Invalid ID in " & ctrl.Name, "ERROR", , "Validation", "GetID"
            GetID = 0
End Function

Private Function GetColumn(ctrl As Control, index As Integer) As String
    On Error Resume Next
    GetColumn = Left(Nz(ctrl.Column(index), IIf(index = 0, 10, 50)))
End Function

Private Sub HandleMissingField(field As Variant)
    On Error Goto ErrorHandler
        Dim ctrl As Control
        Set ctrl = field(0)

        Select Case True
         Case TypeOf ctrl Is ListBox
            MsgBox "Select item in " & field(1), vbExclamation
         Case TypeOf ctrl Is CheckBox
            MsgBox "Check " & field(1), vbExclamation
         Case field(1) = "Range"
            MsgBox "Range format: 1-9999", vbExclamation
            Me.RangeTxt.SetFocus
         Exit Sub
         Case Else
            MsgBox "Required: " & field(1), vbExclamation
        End Select

        ctrl.SetFocus
     Exit Sub

 ErrorHandler:
        LogError "Error handling missing field", "CRITICAL", Err.Number, "Validation", "HandleMissingField"
End Sub

Private Function GetRequiredFieldsCollection() As Collection
    On Error Goto ErrorHandler
        Dim col As New Collection

        ' Core fields
        col.Add Array(Me.BasePrice, "Base Price")
        col.Add Array(Me.EstimatedTotaltxt, "Estimated Total")
        col.Add Array(Me.Airport, "Airport")
        col.Add Array(Me.Vendor, "Vendor")
        col.Add Array(Me.Destination, "Destination")
        col.Add Array(Me.FlightType, "Flight Type")
        col.Add Array(Me.RampName, "Ramp Name")
        col.Add Array(Me.Currency, "Currency")
        col.Add Array(Me.Unit, "Unit")
        col.Add Array(Me.STC, "Subject To Change Checkbox")
        col.Add Array(Me.ChangeCyclebx, "Change Cycle")
        col.Add Array(Me.FuelType, "Fuel Type")
        col.Add Array(Me.Validity, "Validity")
        col.Add Array(Me.Notes, "Notes")
        col.Add Array(Me.FuelingMethodTxt, "Fueling Method")
        ' Conditional field
        If Me.OperatorCbx.value Then
            col.Add Array(Me.Nationality, "Nationality")
            col.Add Array(Me.Operator, "Operator")
        End If

        Set GetRequiredFieldsCollection = col
     Exit Function

 ErrorHandler:
        LogError "Field config error", "CRITICAL", Err.Number, "Validation", "GetRequiredFieldsCollection"
        Set GetRequiredFieldsCollection = New Collection
End Function

Private Function IsFieldMissing(field As Variant) As Boolean
    On Error Goto ErrorHandler
        Dim ctrl As Control
        Set ctrl = field(0)

        Select Case True
         Case TypeOf ctrl Is ListBox
            IsFieldMissing = (ctrl.ItemsSelected.Count = 0)
         Case TypeOf ctrl Is CheckBox
            IsFieldMissing = (Nz(ctrl.value, False) = False)
         Case Else
            IsFieldMissing = (Nz(ctrl.value, "") = "")
        End Select
     Exit Function

 ErrorHandler:
        LogError "Missing field check failed", "CRITICAL", Err.Number, "Validation", "IsFieldMissing"
        IsFieldMissing = True
End Function

'=======================================================
' Transaction Handler
'=======================================================
Private Sub AddFuelPriceBtn_Click()
    On Error Goto ErrorHandler
        Const PROC_NAME As String = "AddFuelPriceBtn_Click"

        Dim ws As DAO.Workspace
        Set ws = DBEngine.Workspaces(0)

        ' Validate first
        Dim validationResult As String
        validationResult = ValidateRequiredFields()
        If validationResult <> "" Then
            LogError "Validation failed: " & validationResult, "ERROR", , "Form", PROC_NAME
         Exit Sub
        End If

        ' Execute transaction
        ws.BeginTrans
        If Add_DropDowns() And Add_FuelPriceDetails() Then
            ws.CommitTrans
            MsgBox "Success!", vbInformation
            DoCmd.Requery
        Else
            ws.Rollback
            MsgBox "Operation failed", vbCritical
        End If

     Exit Sub

 ErrorHandler:
        LogError "Critical error: " & Err.Description, "CRITICAL", Err.Number, "Form", PROC_NAME
        If Not ws Is Nothing Then ws.Rollback
            MsgBox "System error occurred", vbCritical
End Sub
