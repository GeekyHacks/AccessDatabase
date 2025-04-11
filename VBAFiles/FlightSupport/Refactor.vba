Option Compare Database
Option Explicit
Private Sub cmdGenerate_Click()
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Dim randomRefValue As Integer
        Dim RefNo As String
        Dim isUnique As Boolean

        Do
            ' Generate a random number between 1 And 9999
            randomRefValue = Int((9999 - 1 + 1) * Rnd + 1)

            ' Generate the reference number in the New format (e.g., TAS705525)
            RefNo = "TAS0" & Format(randomRefValue, "0000") & Format(Date, "yy")

            ' Check If the generated reference number already exists in the FlightSupport_FlightRequestsT table
            Set rst = CurrentDb.OpenRecordset("Select RefNo FROM FlightSupport_FlightRequestsT WHERE RefNo = '" & RefNo & "' Or RefNo = 'TAS-" & Format(randomRefValue, "0000") & "/" & Format(Date, "yy") & "' Or RefNo = 'TAS-" & Format(randomRefValue, "0000") & "_" & Format(Date, "yy") & "'", dbOpenSnapshot)

            If rst.EOF Then
                ' If the generated reference number does Not exist, assign it To the form field
                Me.RefNo.value = RefNo

                ' Set the flag To indicate that a unique reference number has been generated
                isUnique = True
            End If

            rst.Close
            Set rst = Nothing
        Loop Until isUnique

     Exit Sub

 ErrorHandler:
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub

Private Sub newRequestBtn_Click()
    ' Clear the form fields To prepare For a New record
    ' Set the default value For the RefNo field

    If IsNull(Me.RefNo) Or Me.RefNo = "" Then
        cmdGenerate_Click
    End If
    ' Set focus To the first field For data entry
    Me.ReqNo.SetFocus
End Sub

Private Sub AddFlightBtn_Click()
    newRequestBtn_Click
    ' Check If all the required fields are filled
    If IsNull(Me.RefNo.value) Or IsNull(Me.ReqDate.value) Or IsNull(Me.OPSDate.value) Or IsNull(Me.PFL.value) Or IsNull(Me.InitialRequest.value) Or IsNull(Me.Customer.value) Or IsNull(Me.Customer_IDtxt.value) Or IsNull(Me.ReqNo.value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    Else

        If Me.OPSDate.value = Me.ReqDate.value Then
            Me.Urgent.value = True
        End If
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim currentTimestamp As Variant

        Set db = CurrentDb
        Set rs = db.OpenRecordset("FlightSupport_FlightRequestsT_V13", dbOpenDynaset)
        On Error Goto ErrorHandler
            rs.AddNew

            ' Assign values To the fields in the table
            With rs
                .Fields("RefNo").value = Me.RefNo.value
                .Fields("ReqNo").value = Me.ReqNo.value
                .Fields("ReqDate").value = Me.ReqDate.value
                .Fields("OPSDate").value = Me.OPSDate.value
                .Fields("PFL").value = Me.PFL.value
                .Fields("ScheduleUpdates").value = Me.ScheduleUpdates.value
                .Fields("FuelRequest").value = Me.FuelRequestCb.value
                .Fields("RequestStatus").value = Me.RequestStatus.value
                .Fields("Notes").value = Me.Notes.value
                .Fields("InitialRequest").value = Me.InitialRequest.value
                .Fields("OVFPermit").value = Me.OVFPermitCbx.value
                .Fields("Urgent").value = Me.Urgent.value
                .Fields("Catering").value = Me.CateringCbx.value
                .Fields("Hotel").value = Me.HotelCbx.value
                .Fields("Transportation").value = Me.TransportationCbx.value
                .Fields("LandingPermit").value = Me.LandingPermitCbx.value
                .Fields("GH").value = Me.GHCbx.value
                .Fields("CustomerName").value = Me.Customer.Column(1)
                .Fields("Customer_ID").value = Me.Customer_IDtxt.value

                ' Store the current timestamp For the record
                currentTimestamp = Now
                .Fields("LastModifiedTimestamp").value = currentTimestamp
                .Update
            End With

            ' Attempt To update the record
            'rs.Update

            ' Display a message indicating successful addition
            MsgBox "Request added successfully, Now send request email To the vendor"

            rs.Close
            Set rs = Nothing
            Set db = Nothing

            Me.Undo


            ' Display the last added record
            Dim MaxID As Long
            MaxID = DMax("ID", "FlightSupport_FlightRequestsT")
            Me.RecordSource = "Select CustomerName, Customer_ID, RefNo, ReqNo, ReqDate, OPSDate, PFL, ScheduleUpdates, FuelRequest, FlightSupport, RequestStatus, Notes, InitialRequest, Urgent, GH, OVFPermit, Catering, Hotel, Transportation, LandingPermit FROM FlightSupport_FlightRequestsT WHERE ID = " & MaxID & ";"
            Me.Requery

            ' Clean up

         Exit Sub
        End If

 ErrorHandler:
        Dim ErrorMessage As String
        ErrorMessage = "An error occurred While adding the request:" & vbNewLine
        ErrorMessage = ErrorMessage & "Error Number: " & Err.Number & vbNewLine
        ErrorMessage = ErrorMessage & "Error Description: " & Err.Description & vbNewLine
        MsgBox ErrorMessage, vbExclamation, "Error"

        If Not rs Is Nothing Then
            rs.Close
            Set rs = Nothing
            Set db = Nothing

            Me.Undo
            MsgBox "Failed To add Request successfully"
        End If

        Err.Clear
End Sub

Private Sub UpdateFlightBtn_Click()
    ' Check If ID field is present And valid
    If IsNull(Me.FlightRequestID.value) Or Me.FlightRequestID.value = 0 Then
        MsgBox "No record selected For update. Please Select a record first.", vbExclamation, "No Record Selected"
     Exit Sub
    End If

    ' Check If all required fields are filled
    If IsNull(Me.RefNo.value) Or IsNull(Me.ReqDate.value) Or IsNull(Me.OPSDate.value) Or IsNull(Me.PFL.value) Or _
        IsNull(Me.InitialRequest.value) Or IsNull(Me.Customer.value) Or IsNull(Me.Customer_IDtxt.value) Or IsNull(Me.ReqNo.value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    End If

    ' Set Urgent flag If OPSDate equals ReqDate
    If Me.OPSDate.value = Me.ReqDate.value Then
        Me.Urgent.value = True
    End If

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim currentTimestamp As Variant

    Set db = CurrentDb
    Set rs = db.OpenRecordset("Select * FROM FlightSupport_FlightRequestsT_V13 WHERE ID = " & Me.FlightRequestID.value, dbOpenDynaset)

    On Error Goto ErrorHandler

        If Not rs.EOF Then
            rs.Edit

            ' Update all fields
            With rs
                .Fields("RefNo").value = Me.RefNo.value
                .Fields("ReqNo").value = Me.ReqNo.value
                .Fields("ReqDate").value = Me.ReqDate.value
                .Fields("OPSDate").value = Me.OPSDate.value
                .Fields("PFL").value = Me.PFL.value
                .Fields("ScheduleUpdates").value = Me.ScheduleUpdates.value
                .Fields("FuelRequest").value = Me.FuelRequestCb.value
                .Fields("RequestStatus").value = Me.RequestStatus.value
                .Fields("Notes").value = Me.Notes.value
                .Fields("InitialRequest").value = Me.InitialRequest.value
                .Fields("OVFPermit").value = Me.OVFPermitCbx.value
                .Fields("Urgent").value = Me.Urgent.value
                .Fields("Catering").value = Me.CateringCbx.value
                .Fields("Hotel").value = Me.HotelCbx.value
                .Fields("Transportation").value = Me.TransportationCbx.value
                .Fields("LandingPermit").value = Me.LandingPermitCbx.value
                .Fields("GH").value = Me.GHCbx.value
                .Fields("CustomerName").value = Me.Customer.Column(1)
                .Fields("Customer_ID").value = Me.Customer_IDtxt.value

                ' Update the timestamp
                currentTimestamp = Now
                .Fields("LastModifiedTimestamp").value = currentTimestamp

                .Update
            End With

            MsgBox "Record updated successfully.", vbInformation, "Update Successful"

            ' Refresh the form To show updated data
            Me.Requery
        Else
            MsgBox "Record Not found.", vbExclamation, "Update Failed"
        End If

        rs.Close
        Set rs = Nothing
        Set db = Nothing

     Exit Sub

 ErrorHandler:
        Dim ErrorMessage As String
        ErrorMessage = "An error occurred While updating the record:" & vbNewLine
        ErrorMessage = ErrorMessage & "Error Number: " & Err.Number & vbNewLine
        ErrorMessage = ErrorMessage & "Error Description: " & Err.Description & vbNewLine
        MsgBox ErrorMessage, vbExclamation, "Error"

        If Not rs Is Nothing Then
            rs.Close
            Set rs = Nothing
        End If

        Set db = Nothing
        Err.Clear
End Sub

Private Sub Customer_AfterUpdate()
    Dim selectedValue As Long
    selectedValue = Me.Customer.Column(0) ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("PartiesT_V13", dbOpenSnapshot)

        rs.FindFirst "CorpID = " & selectedValue & ""
        If Not rs.NoMatch Then
            Me.Customer_IDtxt.value = rs("CorpID").value

        Else
            Me.Customer_IDtxt.value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.Customer_IDtxt.value = Null

    End If
    'Me.Requery
 Exit Sub
End Sub

Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim validFileName As String

    ' List of invalid characters in file names
    invalidChars = "\/:*?""<>|"

    ' Replace invalid characters With an underscore
    validFileName = fileName
    For i = 1 To Len(invalidChars)
        validFileName = Replace(validFileName, Mid(invalidChars, i, 1), "_")
    Next i

    ' Remove leading And trailing spaces
    validFileName = Trim(validFileName)

    ' Return the sanitized file name
    SanitizeFileName = validFileName
End Function

Private Sub FullBriefBtn_Click()
    On Error Goto ErrorHandler ' Enable error handling

        Dim strWhere As String
        Dim IngLen As Long

        ' Build the WHERE clause If RefNo is Not null
        If Not IsNull(Me.RefNo) Then
            strWhere = "([RefNo] Like ""*" & Me.RefNo & "*"")"
        End If

        ' Check If the WHERE clause is empty
        IngLen = Len(strWhere)
        If IngLen <= 0 Then
            MsgBox "No parameter specified. Generating report For all data...", vbInformation, "My Flight App"
            ' Open the report without a filter
            DoCmd.OpenReport "FlightSupport_FullExternalReport_V12", acViewPreview
        Else
            ' Apply the filter And open the report in Preview Mode
            DoCmd.OpenReport "FlightSupport_FullExternalReport_V12", acViewPreview, , strWhere
        End If

     Exit Sub

 ErrorHandler:
        MsgBox "Error: " & Err.Description, vbCritical, "Error"

End Sub

Private Sub ExtBriefBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long
    If Not IsNull(Me.RefNo) Then
        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNo & "*"") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_ExternalReport_New", acViewPreview, , strWhere
    End If
End Sub

Private Sub FlightBriefBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    If Not IsNull(Me.RefNo) Then
        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNo & "*"") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_FullInternalReport_V12", acViewPreview, , strWhere
    End If
End Sub

Private Sub Form_Load()

    Dim MaxID As Long


    ' Get the maximum ID from the table
    MaxID = DMax("ID", "FlightSupport_FlightRequestsT")

    ' Update the RecordSource of the current form
    Me.RecordSource = "Select CustomerName, Customer_ID, RefNo, ReqNo, ReqDate, OPSDate, PFL, ScheduleUpdates, FuelRequest, FlightSupport, RequestStatus, Notes, InitialRequest, Urgent, GH, OVFPermit, Catering, Hotel, Transportation, LandingPermit FROM FlightSupport_FlightRequestsT WHERE ID = " & MaxID & ";"
    Me.RefNo2 = Me.RefNo

End Sub

Private Sub OPSDate_Click()
    InputDateField OPSDate, "Select a date To use this on your form"
End Sub

Public Sub TriggerAfterUpdate()
    Call RefNo2_AfterUpdate
End Sub
Private Sub RefNo2_AfterUpdate()
    Dim SelectedRefNo As Variant
    Dim SectorNoCmbQ As String
    Dim CurrentForm As Form

    ' Get the selected RefNo value
    SelectedRefNo = Me.RefNo2.value

    ' Reference the subform within the main form
    Set CurrentForm = Forms("FlightSupport_LASTFORM_New_V1MainFlightSupportF_V13").Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form

    ' Set the record source based on the selected RefNo
    Me.RecordSource = "Select CustomerName, Customer_ID, RefNo, ReqNo, ReqDate, OPSDate, PFL, ScheduleUpdates, FuelRequest, FlightSupport, RequestStatus, Notes, InitialRequest, Urgent, GH, OVFPermit, Catering, Hotel, Transportation, LandingPermit FROM FlightSupport_FlightRequestsT WHERE RefNo = '" & SelectedRefNo & "';"

    ' Construct the query For SectorNoCmb
    SectorNoCmbQ = "Select DISTINCT SectorNo, Sect_ID FROM FlightSupport_SectorsT_New_V12 WHERE FlightRefNo = '" & SelectedRefNo & "' And SectorStatus = 'Active' ORDER BY SectorNo ASC;"

    If Not IsNull(SelectedRefNo) Then
        On Error Resume Next
        If CurrentForm.Form.Controls("SectorRefNo").ControlType <> acEmpty Then
            ' Update controls in the nested subform With the selected RefNo
            CurrentForm.Controls("RequestRefNo").value = SelectedRefNo
            CurrentForm.Controls("SectorRefNo").value = SelectedRefNo
            CurrentForm.Controls("SectorNoCmb").RowSource = SectorNoCmbQ
            ' MsgBox CurrentForm.Controls("SectorNoCmb").RowSource

        Else
            CurrentForm.Controls("RequestRefNo").value = SelectedRefNo
            CurrentForm.Controls("SectorNoCmb").RowSource = SectorNoCmbQ
            CurrentForm.Requery
            MsgBox "The control is Not loaded."
        End If

        On Error Goto 0
            ' Reset dropdowns, reselect, And apply filters
            Reset_DropDowns
            RefNo_reselect
            activeFilter

            ' Requery And refresh the form
            Me.Requery
            Me.Refresh
        Else
         Exit Sub
        End If
End Sub
Private Sub RefNo_reselect()
    Dim SelectedRefNo As Variant
    Dim SectorNoCmbQ As String
    Dim CurrentForm As Form

    SelectedRefNo = Me.RefNo2.value
    Set CurrentForm = Forms("MainFlightSupportF_V13").Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form

    Me.RecordSource = "Select CustomerName, Customer_ID, RefNo,ReqNo,ReqDate,OPSDate,PFL,ScheduleUpdates,FuelRequest,FlightSupport,RequestStatus,Notes,InitialRequest,Urgent, GH, OVFPermit, Catering, Hotel, Transportation, LandingPermit FROM FlightSupport_FlightRequestsT WHERE RefNo = '" & SelectedRefNo & "';"
    SectorNoCmbQ = "Select DISTINCT SectorNo, Sect_ID FROM FlightSupport_SectorsT_New_V12 WHERE FlightRefNo = '" & SelectedRefNo & "' And SectorStatus = 'Active' ORDER BY SectorNo ASC;"
    On Error Resume Next
    If CurrentForm.Form.Controls("ConcierageSectorCmb").ControlType <> acEmpty Or CurrentForm.Form.Controls("PermitSectorsCmb").ControlType <> acEmpty Or CurrentForm.Form.Controls("HandlingSectorCmb").ControlType <> acEmpty Then
        CurrentForm.Form.Controls("PermitSectorsCmb").RowSource = SectorNoCmbQ
        CurrentForm.Form.Controls("HandlingSectorCmb").RowSource = SectorNoCmbQ
        CurrentForm.Form.Controls("ConcierageSectorCmb").RowSource = SectorNoCmbQ
        CurrentForm.Form.Controls("FlightSupport_ServicesTabs").Controls("BriefsTabs").Form.Controls("Sect_ID").RowSource = SectorNoCmbQ

    Else
        MsgBox "The control is Not loaded."
    End If

    On Error Goto 0 ' Disable error handling To return To normal error behavior
        Me.Requery
        Me.Refresh
End Sub
Private Sub Reset_DropDowns()
    Dim CurrentForm As Form
    Set CurrentForm = Forms("MainFlightSupportF_V13").Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form

    On Error Resume Next
    If CurrentForm.Form.Controls("ConcierageSectorCmb").ControlType <> acEmpty Or CurrentForm.Form.Controls("PermitSectorsCmb").ControlType <> acEmpty Or CurrentForm.Form.Controls("HandlingSectorCmb").ControlType <> acEmpty Then
        CurrentForm.Form.Controls("PermitSectorsCmb").value = Null
        CurrentForm.Form.Controls("HandlingSectorCmb").value = Null
        CurrentForm.Form.Controls("ConcierageSectorCmb").value = Null
        CurrentForm.Form.Controls("FlightSupport_ServicesTabs").Controls("BriefsTabs").Form.Controls("Sect_ID").value = Null
    Else
        CurrentForm.Form.Controls("SectorNoCmb").value = Null
        MsgBox "The control is Not loaded."
    End If

    On Error Goto 0 ' Disable error handling To return To normal error behavior

End Sub

Private Sub activeFilter()
    Dim SelectedRefNo As Variant
    Dim query As String
    Dim CurrentForm As Form

    Set CurrentForm = Forms("MainFlightSupportF_V13").Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form
    SelectedRefNo = Me.RefNo2.value

    On Error Resume Next
    If CurrentForm.Form.Controls("ConcierageSectorCmb").ControlType <> acEmpty Or CurrentForm.Form.Controls("PermitSectorsCmb").ControlType <> acEmpty Or CurrentForm.Form.Controls("HandlingSectorCmb").ControlType <> acEmpty Then
        CurrentForm.Form.Controls("FlightSupport_UpdateSectorsF_New").Form.Filter = "RequestRefNo = '" & SelectedRefNo & "'"
        CurrentForm.Form.Controls("FlightSupport_UpdateSectorsF_New").Form.FilterOn = True
        CurrentForm.Form.Controls("FlightSupport_SectorsListF_New").Form.Filter = "RequestRefNo = '" & SelectedRefNo & "'"
        CurrentForm.Form.Controls("FlightSupport_SectorsListF_New").Form.FilterOn = True
    Else
        MsgBox "The control is Not loaded."
    End If
End Sub
Public Sub ReqDate_Click()
    InputDateField ReqDate, "Select a date To use this on your form"
End Sub


Private Sub ResetBtn_Click()
    On Error Goto ErrorHandler

        Me.Undo
        DoCmd.CancelEvent
        DoCmd.GoToRecord , , acNewRec

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & Err.Description, vbExclamation
End Sub

Private Sub LoadLastRequestBtn_Click()

    Dim MaxID As Long


    ' Get the maximum ID from the table
    MaxID = DMax("ID", "FlightSupport_FlightRequestsT")

    ' Update the RecordSource of the current form
    Me.RecordSource = "Select CustomerName, Customer_ID, RefNo, ReqNo, ReqDate, OPSDate, PFL, ScheduleUpdates, FuelRequest, FlightSupport, RequestStatus, Notes, InitialRequest, Urgent, GH, OVFPermit, Catering, Hotel, Transportation, LandingPermit FROM FlightSupport_FlightRequestsT WHERE ID = " & MaxID & ";"
    Me.RefNo2 = Me.RefNo

    Dim MainForm As Form
    Dim SubFormControl As Control
    Dim SubForm As Form

    ' Reference the main form
    Set MainForm = Forms("MainFlightSupportF_V13")

    ' Reference the subform control on the main form
    Set SubFormControl = MainForm.Controls("FlightSupport_AddUpdateRequestF")

    ' Reference the form inside the subform control
    Set SubForm = SubFormControl.Form

    Set FinalSubForm = SubForm.Controls("Add_UpdateRequestF")


    ' Reference the control on the subform
    FinalSubForm.Controls("RequestRefNo").value = Me.RefNo
    FinalSubForm.Controls("SectorRefNo").value = Me.RefNo
    ' Requery the current form
    Me.Requery

End Sub

Private Sub ScheduleUpdates_Change()
    Me.OPSDate.value = Format(Now, "dd-MM-yyyy")
End Sub
/////////////////////////////////////////////////////////////////////////////////////////
Option Compare Database
Option Explicit

'=============== MODULE DECLARATIONS ===============
Private Const PRIMARY_TABLE As String = "FlightSupport_FlightRequestsT_V13"
Private m_lngCurrentID As Long  'Stores current record ID
Private m_bIsLoading As Boolean 'Prevents updates during load

'=============== FORM EVENTS ===============
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    'Completely unbound form - no RecordSource needed
    Me.RecordSource = ""
    
    'Optional: Load last record
    LoadLastRecord
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vbCritical
End Sub

'=============== FIELD UPDATE HANDLERS ===============
Private Sub ReqNo_AfterUpdate()
    If Not m_bIsLoading Then UpdateField "ReqNo", Me.ReqNo.Value
End Sub

Private Sub ReqDate_AfterUpdate()
    If Not m_bIsLoading Then 
        UpdateField "ReqDate", Me.ReqDate.Value
        'Auto-update Urgent flag
        Me.Urgent.Value = (Me.OPSDate.Value = Me.ReqDate.Value)
        UpdateField "Urgent", Me.Urgent.Value
    End If
End Sub

'... Add similar AfterUpdate handlers for all editable fields ...

'=============== CORE FUNCTIONS ===============
Private Sub UpdateField(sFieldName As String, vValue As Variant)
    On Error GoTo ErrorHandler
    
    'Only update if we have a loaded record
    If m_lngCurrentID = 0 Then Exit Sub
    
    Dim strSQL As String
    strSQL = "UPDATE " & PRIMARY_TABLE & " SET " & _
             sFieldName & " = ?, LastModifiedTimestamp = Now() " & _
             "WHERE ID = ?"
             
    CurrentDb.Execute strSQL, dbFailOnError, Array(vValue, m_lngCurrentID)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating " & sFieldName & ": " & Err.Description, vbExclamation
End Sub

Private Sub LoadLastRecord()
    On Error GoTo ErrorHandler
    Dim rst As DAO.Recordset
    Dim MaxID As Long
    
    m_bIsLoading = True 'Prevent update triggers
    
    MaxID = DMax("ID", PRIMARY_TABLE)
    If MaxID > 0 Then
        Set rst = CurrentDb.OpenRecordset( _
            "SELECT * FROM " & PRIMARY_TABLE & " WHERE ID = " & MaxID, _
            dbOpenSnapshot)
        
        If Not rst.EOF Then
            m_lngCurrentID = MaxID
            'Load values into unbound controls
            Me.RefNo.Value = rst!RefNo
            Me.ReqNo.Value = rst!ReqNo
            '... load all other fields ...
        End If
        rst.Close
    End If
    
    m_bIsLoading = False
    Exit Sub
    
ErrorHandler:
    m_bIsLoading = False
    If Not rst Is Nothing Then rst.Close
    MsgBox "Error loading record: " & Err.Description, vbExclamation
End Sub

Private Sub ClearFormControls()
    'Clear all unbound controls
    Me.RefNo.Value = Null
    Me.ReqNo.Value = Null
    '... clear all other fields ...
    m_lngCurrentID = 0 'Reset current record
End Sub

Private Sub AddFlightBtn_Click()
    On Error GoTo ErrorHandler
    
    If Not ValidateForm() Then Exit Sub
    
    Dim strSQL As String
    strSQL = "INSERT INTO " & PRIMARY_TABLE & " (RefNo, ReqNo, ...) " & _
             "VALUES (?, ?, ...)"
    
    CurrentDb.Execute strSQL, dbFailOnError, Array( _
        Me.RefNo.Value, Me.ReqNo.Value, ... 'All field values
    )
    
    'Reload the new record
    LoadRecordByID DMax("ID", PRIMARY_TABLE)
    
    MsgBox "Record added successfully", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding record: " & Err.Description, vbCritical
End Sub


