Option Compare Database
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (Byval dwMilliseconds As LongPtr)
    #End If

    Private Const FORM_NAME As String = "Add_UpdateRequestF"
    Private Const PRIMARY_TABLE As String = "FlightSupport_FlightRequestsT"
    Private Const BATCH_TIMEOUT_MS As Long = 2000 ' 2 seconds
    Private Const MAX_BATCH_SIZE As Long = 5 ' Max updates before auto-commit

    Private Type tPendingUpdate
    FieldName As String
    FieldValue As Variant
    End Type

    Private m_lngCurrentID As Long
    Private m_bIsLoading As Boolean
    Private m_PendingUpdates() As tPendingUpdate
    Private m_UpdateCount As Long
    Private m_LastUpdateTime As Date

Private Sub Form_Open(Cancel As Integer)
    'Me.Visible = True
    'DoEvents
End Sub

Private Sub Form_Load()
    On Error Goto ErrorHandler

        ' Initialize As unbound form
        Me.RecordSource = ""
        Me.RecordsetType = 0 ' Form (Not Snapshot)

        ' Initialize batch update system
        ReDim m_PendingUpdates(MAX_BATCH_SIZE - 1)
        m_UpdateCount = 0
        Me.TimerInterval = 500 ' Check every 0.5 seconds

        ' Set loading flag
        m_bIsLoading = True

        ' Load the last record (won't clear controls first)
        LoadLastRecord

        ' Reset loading flag
        m_bIsLoading = False

        ' Ensure form is visible
        Me.Visible = True
        DoEvents

     Exit Sub

 ErrorHandler:
        m_bIsLoading = False
        LogError Err.Description, "Form_Load", Err.Number, FORM_NAME, "Form_Load"
        MsgBox "Error loading form: " & Err.Description, vbExclamation
        Me.Visible = True
End Sub

Private Sub Form_Timer()
    ' Commit any pending updates If timeout reached
    If m_UpdateCount > 0 And DateDiff("s", m_LastUpdateTime, Now) * 1000 > BATCH_TIMEOUT_MS Then
        CommitPendingUpdates
    End If
End Sub

'=======================================================
' CORE FUNCTIONALITY
'=======================================================

Private Sub cmdGenerate_Click()
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Dim randomRefValue As Integer
        Dim RefNo As String
        Dim isUnique As Boolean
        Dim startTime As Variant: startTime = Now

        Do
            ' Generate random number between 1 And 9999
            randomRefValue = Int((9999 - 1 + 1) * Rnd + 1)

            ' Create reference number format
            RefNo = "TAS0" & Format(randomRefValue, "0000") & Format(Date, "yy")

            ' Check uniqueness With minimal query
            Set rst = CurrentDb.OpenRecordset( _
            "Select 1 FROM " & PRIMARY_TABLE & " WHERE RefNo = '" & RefNo & "'", _
            dbOpenSnapshot)

            isUnique = rst.EOF
            rst.Close

        Loop Until isUnique Or DateDiff("s", startTime, Now) > 5 ' 5 second timeout

        If isUnique Then
            Me.RefNo.value = RefNo
        Else
            MsgBox "Could Not generate unique reference number. Please try again.", vbExclamation
        End If

     Exit Sub

 ErrorHandler:
        LogError Err.Description, "cmdGenerate_Click", Err.Number, FORM_NAME, "cmdGenerate_Click", "RefNo: " & RefNo
        If Not rst Is Nothing Then rst.Close
            Set rst = Nothing
            MsgBox "Error generating reference number: " & Err.Description, vbCritical
End Sub

Private Sub newRequestBtn_Click()
    On Error Goto ErrorHandler

        ' Commit any pending updates before clearing
        If m_UpdateCount > 0 Then
            CommitPendingUpdates
        End If

        ClearFormControls
        If IsNull(Me.RefNo) Or Me.RefNo = "" Then cmdGenerate_Click
            Me.ReqNo.SetFocus

         Exit Sub

 ErrorHandler:
            LogError Err.Description, "newRequestBtn_Click", Err.Number, FORM_NAME, "newRequestBtn_Click"
            MsgBox "Error preparing New request: " & Err.Description, vbExclamation
End Sub

Private Sub AddFlightBtn_Click()
    On Error Goto ErrorHandler

        ' Commit any pending updates first
        If m_UpdateCount > 0 Then
            CommitPendingUpdates
        End If

        ' Validate before any database operations
        If Not ValidateForm() Then Exit Sub

            ' Declare transaction variables
            Dim ws As DAO.Workspace
            Dim qdf As DAO.QueryDef
            Dim strSQL As String
            Dim lngID As Long

            ' Set urgent flag If same date
            Me.Urgent.value = (Me.OPSDate.value = Me.ReqDate.value)

            ' Begin transaction
            Set ws = DBEngine.Workspaces(0)
            ws.BeginTrans

            ' Build parameterized query
            strSQL = "PARAMETERS [pRefNo] Text(20), [pReqNo] Text(20), [pReqDate] DateTime, " & _
            "[pOPSDate] DateTime, [pPFL] Text(20), [pScheduleUpdates] Memo, " & _
            "[pFuelRequest] Bit, [pRequestStatus] Text(20), [pNotes] Memo, " & _
            "[pInitialRequest] Memo, [pOVFPermit] Bit, [pUrgent] Bit, [pCatering] Bit, " & _
            "[pHotel] Bit, [pTransportation] Bit, [pLandingPermit] Bit, [pGH] Bit, " & _
            "[pCustomerName] Text(100), [pCustomer_ID] Long; " & _
            "INSERT INTO " & PRIMARY_TABLE & " (" & _
            "RefNo, ReqNo, ReqDate, OPSDate, PFL, ScheduleUpdates, " & _
            "FuelRequest, RequestStatus, Notes, InitialRequest, " & _
            "OVFPermit, Urgent, Catering, Hotel, Transportation, " & _
            "LandingPermit, GH, CustomerName, Customer_ID, LastModifiedTimestamp" & _
            ") VALUES (" & _
            "[pRefNo], [pReqNo], [pReqDate], [pOPSDate], [pPFL], [pScheduleUpdates], " & _
            "[pFuelRequest], [pRequestStatus], [pNotes], [pInitialRequest], " & _
            "[pOVFPermit], [pUrgent], [pCatering], [pHotel], [pTransportation], " & _
            "[pLandingPermit], [pGH], [pCustomerName], [pCustomer_ID], Now())"

            ' Execute the insert
            Set qdf = CurrentDb.CreateQueryDef("", strSQL)
            With qdf
                .Parameters("pRefNo") = Me.RefNo.value
                .Parameters("pReqNo") = Me.ReqNo.value
                .Parameters("pReqDate") = Me.ReqDate.value
                .Parameters("pOPSDate") = Me.OPSDate.value
                .Parameters("pPFL") = Nz(Me.PFL.value, "")
                .Parameters("pScheduleUpdates") = Nz(Me.ScheduleUpdates.value, "")
                .Parameters("pFuelRequest") = Me.FuelRequestCb.value
                .Parameters("pRequestStatus") = Nz(Me.RequestStatus.value, "")
                .Parameters("pNotes") = Nz(Me.Notes.value, "")
                .Parameters("pInitialRequest") = Nz(Me.InitialRequest.value, "")
                .Parameters("pOVFPermit") = Me.OVFPermitCbx.value
                .Parameters("pUrgent") = Me.Urgent.value
                .Parameters("pCatering") = Me.CateringCbx.value
                .Parameters("pHotel") = Me.HotelCbx.value
                .Parameters("pTransportation") = Me.TransportationCbx.value
                .Parameters("pLandingPermit") = Me.LandingPermitCbx.value
                .Parameters("pGH") = Me.GHCbx.value
                .Parameters("pCustomerName") = Me.Customer.Column(1)
                .Parameters("pCustomer_ID") = Me.Customer_IDtxt.value

                .Execute dbFailOnError
            End With

            ' Get the New ID
            lngID = DMax("ID", PRIMARY_TABLE)

            ' Load the New record - this also does validation
            LoadRecordByID lngID

            ' Commit the transaction If everything succeeded
            ws.CommitTrans

            ' Show success message
            MsgBox "Request added successfully", vbInformation

         Exit Sub

 ErrorHandler:
            ' Roll back transaction If it was started
            If Not ws Is Nothing Then
                If ws.Transactions > 0 Then ws.Rollback
                End If

                ' Log the error
                LogError Err.Description, "AddFlightBtn_Click", Err.Number, FORM_NAME, "AddFlightBtn_Click", "SQL: " & strSQL

                ' Show error message
                MsgBox "Error adding request: " & Err.Description, vbCritical

                ' Clean up objects
                If Not qdf Is Nothing Then Set qdf = Nothing
                    If Not ws Is Nothing Then Set ws = Nothing
End Sub

'=======================================================
' FIELD UPDATE HANDLERS
'=======================================================

Private Sub ReqNo_AfterUpdate()
    UpdateField "ReqNo", Me.ReqNo.value
End Sub

Private Sub ReqDate_AfterUpdate()
    UpdateField "ReqDate", Me.ReqDate.value
    ' Auto-update Urgent flag If dates match
    If Not m_bIsLoading Then
        Me.Urgent.value = (Me.OPSDate.value = Me.ReqDate.value)
        UpdateField "Urgent", Me.Urgent.value
    End If
End Sub

Private Sub OPSDate_AfterUpdate()
    UpdateField "OPSDate", Me.OPSDate.value
    ' Auto-update Urgent flag If dates match
    If Not m_bIsLoading Then
        Me.Urgent.value = (Me.OPSDate.value = Me.ReqDate.value)
        UpdateField "Urgent", Me.Urgent.value
    End If
End Sub

Private Sub PFL_AfterUpdate()
    UpdateField "PFL", Me.PFL.value
End Sub

Private Sub ScheduleUpdates_AfterUpdate()
    UpdateField "ScheduleUpdates", Me.ScheduleUpdates.value
End Sub

Private Sub FuelRequestCb_AfterUpdate()
    UpdateField "FuelRequest", Me.FuelRequestCb.value
End Sub

Private Sub RequestStatus_AfterUpdate()
    UpdateField "RequestStatus", Me.RequestStatus.value
End Sub

Private Sub Notes_AfterUpdate()
    UpdateField "Notes", Me.Notes.value
End Sub

Private Sub InitialRequest_AfterUpdate()
    UpdateField "InitialRequest", Me.InitialRequest.value
End Sub

Private Sub OVFPermitCbx_AfterUpdate()
    UpdateField "OVFPermit", Me.OVFPermitCbx.value
End Sub

Private Sub CateringCbx_AfterUpdate()
    UpdateField "Catering", Me.CateringCbx.value
End Sub

Private Sub HotelCbx_AfterUpdate()
    UpdateField "Hotel", Me.HotelCbx.value
End Sub

Private Sub TransportationCbx_AfterUpdate()
    UpdateField "Transportation", Me.TransportationCbx.value
End Sub

Private Sub LandingPermitCbx_AfterUpdate()
    UpdateField "LandingPermit", Me.LandingPermitCbx.value
End Sub

Private Sub GHCbx_AfterUpdate()
    UpdateField "GH", Me.GHCbx.value
End Sub

Private Sub Customer_AfterUpdate()
    On Error Goto ErrorHandler

        ' Commit any pending updates first
        If m_UpdateCount > 0 Then
            CommitPendingUpdates
        End If

        Dim ws As DAO.Workspace
        Dim rs As DAO.Recordset

        ' Begin transaction
        Set ws = DBEngine.Workspaces(0)
        ws.BeginTrans

        If Not IsNull(Me.Customer.value) Then
            Set rs = CurrentDb.OpenRecordset( _
            "Select CorpID FROM PartiesT_V13 WHERE CorpID = " & Me.Customer.Column(0), _
            dbOpenSnapshot)

            If Not rs.EOF Then
                Me.Customer_IDtxt.value = rs!CorpID

                ' Update both fields (batch these together)
                AddPendingUpdate "Customer_ID", rs!CorpID
                AddPendingUpdate "CustomerName", Me.Customer.Column(1)
                CommitPendingUpdates ' Immediate commit For related fields
            Else
                Me.Customer_IDtxt.value = Null
            End If
            rs.Close
        Else
            Me.Customer_IDtxt.value = Null
        End If

        ' Commit transaction
        ws.CommitTrans

     Exit Sub

 ErrorHandler:
        ' Roll back transaction If it was started
        If Not ws Is Nothing Then
            If ws.Transactions > 0 Then ws.Rollback
            End If

            ' Clean up
            If Not rs Is Nothing Then rs.Close
                Set rs = Nothing
                Set ws = Nothing

                ' Log And show error
                LogError Err.Description, "Customer_AfterUpdate", Err.Number, FORM_NAME, "Customer_AfterUpdate"
                MsgBox "Error updating customer ID: " & Err.Description, vbExclamation
End Sub

'=======================================================
' BATCH UPDATE SYSTEM
'=======================================================

Private Sub UpdateField(sFieldName As String, vValue As Variant)
    On Error Goto ErrorHandler

        ' Check If we have a valid record loaded
        If m_lngCurrentID = 0 Then Exit Sub

            ' Visual feedback
            Me.Controls(sFieldName).BackColor = RGB(200, 255, 200) ' Light green
            Me.Controls(sFieldName).FontBold = True
            DoEvents

            ' Add To pending updates
            AddPendingUpdate sFieldName, vValue

            ' Start timer For batch commit
            m_LastUpdateTime = Now

            ' Check If we should commit immediately
            If m_UpdateCount >= MAX_BATCH_SIZE Or DateDiff("s", m_LastUpdateTime, Now) * 1000 > BATCH_TIMEOUT_MS Then
                CommitPendingUpdates
            End If

         Exit Sub

 ErrorHandler:
            Me.Controls(sFieldName).BackColor = vbWhite
            Me.Controls(sFieldName).FontBold = False
            MsgBox "Error queueing update For " & sFieldName & ":" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub AddPendingUpdate(sFieldName As String, vValue As Variant)
    ' Resize array If needed
    If m_UpdateCount >= UBound(m_PendingUpdates) + 1 Then
        ReDim Preserve m_PendingUpdates(UBound(m_PendingUpdates) + MAX_BATCH_SIZE)
    End If

    ' Add the update
    With m_PendingUpdates(m_UpdateCount)
        .FieldName = sFieldName
        .FieldValue = vValue
    End With
    m_UpdateCount = m_UpdateCount + 1

    ' Log the queued update
    Debug.Print "Queued update: " & sFieldName & " = " & CStr(vValue)
End Sub

Private Sub CommitPendingUpdates()
    On Error Goto ErrorHandler

        If m_UpdateCount = 0 Then Exit Sub

            Dim ws As DAO.Workspace
            Dim qdf As DAO.QueryDef
            Dim strSQL As String
            Dim i As Long
            Dim originalTimestamp As Variant

            ' Concurrency check
            originalTimestamp = DLookup("LastModifiedTimestamp", PRIMARY_TABLE, "ID = " & m_lngCurrentID)
            If Not IsNull(originalTimestamp) Then
                If originalTimestamp <> Nz(Me.LastModifiedTag.value, 0) Then
                    MsgBox "Record was modified by another user. Please refresh.", vbExclamation
                    ClearPendingUpdates
                 Exit Sub
                End If
            End If

            ' Begin transaction
            Set ws = DBEngine.Workspaces(0)
            ws.BeginTrans

            ' Build dynamic SQL For all pending updates
            strSQL = "UPDATE " & PRIMARY_TABLE & " Set "

            For i = 0 To m_UpdateCount - 1
                If i > 0 Then strSQL = strSQL & ", "
                    strSQL = strSQL & m_PendingUpdates(i).FieldName & " = [pVal" & i & "]"
                Next i

                strSQL = strSQL & ", LastModifiedTimestamp = Now() WHERE ID = [pID]"

                ' Execute the update
                Set qdf = CurrentDb.CreateQueryDef("", strSQL)

                For i = 0 To m_UpdateCount - 1
                    qdf.Parameters("pVal" & i).value = m_PendingUpdates(i).FieldValue
                Next i

                qdf.Parameters("pID").value = m_lngCurrentID
                qdf.Execute dbFailOnError

                ' Commit transaction
                ws.CommitTrans

                ' Update the last modified tag
                Me.LastModifiedTag.value = DLookup("LastModifiedTimestamp", PRIMARY_TABLE, "ID = " & m_lngCurrentID)

                ' Log the batch update
                Debug.Print "Committed " & m_UpdateCount & " updates at " & Now()

                ' Clear pending updates
                ClearPendingUpdates

                ' Restore control appearances
                For i = 0 To m_UpdateCount - 1
                    On Error Resume Next ' In Case control is no longer available
                    Me.Controls(m_PendingUpdates(i).FieldName).BackColor = vbWhite
                    Me.Controls(m_PendingUpdates(i).FieldName).FontBold = False
                    On Error Goto 0
                    Next i

                 Exit Sub

 ErrorHandler:
                    If Not ws Is Nothing Then
                        If ws.Transactions > 0 Then ws.Rollback
                        End If

                        LogError Err.Description, "CommitPendingUpdates", Err.Number, FORM_NAME, "CommitPendingUpdates", _
                        "Updates: " & m_UpdateCount & ", SQL: " & strSQL

                        MsgBox "Error committing updates: " & Err.Description, vbExclamation

                        If Not qdf Is Nothing Then Set qdf = Nothing
                            If Not ws Is Nothing Then Set ws = Nothing

                                ' Restore control appearances On Error
                                For i = 0 To m_UpdateCount - 1
                                    On Error Resume Next
                                    Me.Controls(m_PendingUpdates(i).FieldName).BackColor = vbWhite
                                    Me.Controls(m_PendingUpdates(i).FieldName).FontBold = False
                                    On Error Goto 0
                                    Next i
End Sub

Private Sub ClearPendingUpdates()
    Erase m_PendingUpdates
    m_UpdateCount = 0
    ReDim m_PendingUpdates(MAX_BATCH_SIZE - 1)
End Sub

'=======================================================
' RECORD MANAGEMENT FUNCTIONS
'=======================================================

Private Sub LoadLastRecord()
    On Error Goto ErrorHandler

        Debug.Print "Attempting To load last record..."
        Dim MaxID As Long
        MaxID = DMax("ID", PRIMARY_TABLE)
        Debug.Print "Max ID found: " & MaxID

        If MaxID > 0 Then
            LoadRecordByID MaxID
        Else
            Debug.Print "No records found in table"
            ' Optional: Show message If no records exist
            MsgBox "No existing records found", vbInformation
        End If

     Exit Sub

 ErrorHandler:
        LogError Err.Description, "LoadLastRecord", Err.Number, FORM_NAME, "LoadLastRecord"
        MsgBox "Error loading last record: " & Err.Description, vbExclamation
End Sub

Private Sub LoadRecordByID(lngID As Long)
    On Error Goto ErrorHandler

        ' Commit any pending updates before loading New record
        If m_UpdateCount > 0 Then
            CommitPendingUpdates
        End If

        m_bIsLoading = True

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset( _
        "Select * FROM " & PRIMARY_TABLE & " WHERE ID = " & lngID, _
        dbOpenSnapshot)

        If Not rst.EOF Then
            ' Store ID For updates
            m_lngCurrentID = lngID

            ' Load values into controls
            Me.RefNo.value = Nz(rst!RefNo, "")
            Me.ReqNo.value = Nz(rst!ReqNo, "")
            Me.ReqDate.value = Nz(rst!ReqDate, "")
            Me.OPSDate.value = Nz(rst!OPSDate, "")
            Me.PFL.value = Nz(rst!PFL, "")
            Me.ScheduleUpdates.value = Nz(rst!ScheduleUpdates, "")
            Me.FuelRequest.value = Nz(rst!FuelRequest, False)
            Me.RequestStatus.value = Nz(rst!RequestStatus, "")
            Me.Notes.value = Nz(rst!Notes, "")
            Me.InitialRequest.value = Nz(rst!InitialRequest, "")
            Me.OVFPermit.value = Nz(rst!OVFPermit, False)
            Me.Urgent.value = Nz(rst!Urgent, False)
            Me.Catering.value = Nz(rst!Catering, False)
            Me.Hotel.value = Nz(rst!Hotel, False)
            Me.Transportation.value = Nz(rst!Transportation, False)
            Me.LandingPermit.value = Nz(rst!LandingPermit, False)
            Me.GH.value = Nz(rst!GH, False)
            Me.Customer.value = Nz(rst!CustomerName, Null)
            Me.Customer_IDtxt.value = Nz(rst!Customer_ID, "")
            Me.LastModifiedTag.value = Nz(rst!LastModifiedTimestamp, 0)
        End If

        rst.Close
        m_bIsLoading = False

        ' Force form refresh
        Me.Repaint
        DoEvents

     Exit Sub

 ErrorHandler:
        m_bIsLoading = False
        If Not rst Is Nothing Then rst.Close
            LogError Err.Description, "LoadRecordByID", Err.Number, FORM_NAME, "LoadRecordByID", "ID: " & lngID
            MsgBox "Error loading record: " & Err.Description, vbExclamation
End Sub

'=======================================================
' FORM MANAGEMENT FUNCTIONS
'=======================================================

Private Sub ClearFormControls()
    On Error Goto ErrorHandler

        ' Commit any pending updates before clearing
        If m_UpdateCount > 0 Then
            CommitPendingUpdates
        End If

        ' Clear all controls
        Me.RefNo.value = Null
        Me.ReqNo.value = Null
        Me.ReqDate.value = Null
        Me.OPSDate.value = Null
        Me.PFL.value = Null
        Me.ScheduleUpdates.value = Null
        Me.FuelRequestCb.value = False
        Me.RequestStatus.value = Null
        Me.Notes.value = Null
        Me.InitialRequest.value = Null
        Me.OVFPermitCbx.value = False
        Me.Urgent.value = False
        Me.CateringCbx.value = False
        Me.HotelCbx.value = False
        Me.TransportationCbx.value = False
        Me.LandingPermitCbx.value = False
        Me.GHCbx.value = False
        Me.Customer.value = Null
        Me.Customer_IDtxt.value = Null

        ' Reset current ID
        m_lngCurrentID = 0

     Exit Sub

 ErrorHandler:
        LogError Err.Description, "ClearFormControls", Err.Number, FORM_NAME, "ClearFormControls"
        MsgBox "Error clearing form controls: " & Err.Description, vbExclamation
End Sub

'=======================================================
' VALIDATION FUNCTIONS
'=======================================================

Private Function ValidateForm() As Boolean
    Dim arrRequired As Variant
    Dim i As Integer
    Dim strMissing As String

    arrRequired = Array("RefNo", "ReqNo", "ReqDate", "OPSDate", "PFL", _
    "InitialRequest", "Customer", "Customer_IDtxt")

    For i = LBound(arrRequired) To UBound(arrRequired)
        If IsNull(Me.Controls(arrRequired(i)).value) Then
            strMissing = strMissing & ", " & arrRequired(i)
        End If
    Next i

    If strMissing <> "" Then
        MsgBox "Missing required fields: " & Mid(strMissing, 3), vbExclamation
        ValidateForm = False
    Else
        ValidateForm = True
    End If
End Function

'=======================================================
' LOGGING SYSTEM
'=======================================================

Private Sub LogError( _
    Byval ErrorMessage As String, _
    Byval ErrorSource As String, _
    Optional Byval ErrorNumber As Long = 0, _
    Optional Byval ModuleName As String = "", _
    Optional Byval ProcedureName As String = "", _
    Optional Byval AdditionalInfo As String = "")

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

        filePath = CurrentProject.Path & "\ErrorLog_FlightRequest.log"
        fileNumber = FreeFile

        Open filePath For Append As #fileNumber
        Print #fileNumber, logMessage
        Close #fileNumber

     Exit Sub

 LogError_Handler:
        ' If logging fails, show message but don't crash
        MsgBox "Logging failed: " & Err.Description & vbCrLf & _
        "Original error: " & ErrorMessage, vbCritical
End Sub

Private Sub LogChange(sFieldName As String, vValue As Variant)
    Dim filePath As String
    Dim fileNumber As Integer
    Dim logMessage As String

    logMessage = "Field Changed: " & sFieldName & vbCrLf & _
    "New Value: " & CStr(vValue) & vbCrLf & _
    "Record ID: " & m_lngCurrentID & vbCrLf & _
    "Timestamp: " & Now() & vbCrLf & _
    String(30, "-") & vbCrLf

    filePath = CurrentProject.Path & "\FieldChanges.log"
    fileNumber = FreeFile

    Open filePath For Append As #fileNumber
    Print #fileNumber, logMessage
    Close #fileNumber

 Exit Sub
End Sub

Private Function GetFieldType(sFieldName As String) As String
    ' Map field names To SQL data types
    Select Case sFieldName
     Case "ReqDate", "OPSDate", "LastModifiedTimestamp"
        GetFieldType = "DateTime"
     Case "FuelRequest", "Urgent", "OVFPermit", "Catering", "Hotel", "Transportation", "LandingPermit", "GH"
        GetFieldType = "Bit"
     Case "Customer_ID"
        GetFieldType = "Long"
     Case "RefNo", "ReqNo", "PFL", "RequestStatus", "CustomerName"
        GetFieldType = "Text(100)"
     Case "ScheduleUpdates", "Notes", "InitialRequest"
        GetFieldType = "Memo"
     Case Else
        GetFieldType = "Text(255)" ' Default
    End Select
End Function