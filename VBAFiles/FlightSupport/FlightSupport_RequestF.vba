Option Compare Database
Private Sub ACRegCb_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType.Value = rs("ACType").Value
            Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator.Value = rs("Operator").Value
            Me.Customer.Value = rs("Customer").Value
        Else
            Me.ACType.Value = Null
            Me.ACMTOW.Value = Null
            Me.Operator.Value = Null
            Me.Customer.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ACType.Value = Null
        Me.ACMTOW.Value = Null
        Me.Operator.Value = Null
        Me.Customer.Value = Null
    End If

 Exit Sub

End Sub
Private Sub cmdGenerate_Click()
    Dim rst As DAO.Recordset
    Dim randomRefValue As Integer
    Dim RefNo As String
    Dim isUnique As Boolean

    Do
        ' Generate a random number between 1 And 999
        randomRefValue = Int((9999 - 1 + 1) * Rnd + 1)

        ' Check If the generated number is greater than Or equal To 176
        If randomRefValue >= 176 Then
            ' Generate the reference number
            RefNo = "TAS-" & Format(randomRefValue, "0000") & "/" & Format(date, "yy")

            ' Check If the generated reference number already exists in the NewFlightRequestsT table
            Set rst = CurrentDb.OpenRecordset("Select RefNo FROM NewFlightRequestsT WHERE RefNo = '" & RefNo & "'", dbOpenSnapshot)

            If rst.EOF Then
                ' If the generated reference number does Not exist, assign it To the form field
                Me.RefNo.Value = RefNo

                ' Set the flag To indicate that a unique reference number has been generated
                isUnique = True
            End If

            rst.Close
            Set rst = Nothing
        End If
    Loop Until isUnique

End Sub
Private Sub AddFlightBtn_Click()
    ' Check If all the required fields are filled
    If IsNull(Me.RefNo.Value) Or IsNull(Me.ReqDate.Value) Or IsNull(Me.OPSDate.Value) Or IsNull(Me.Customer.Value) Or IsNull(Me.Operator.Value) _
        Or IsNull(Me.ACType.Value) Or IsNull(Me.ACMTOW.Value) Or IsNull(Me.ACRegCb.Value) Or IsNull(Me.PFL.Value) Or IsNull(Me.Schedule.Value) Or IsNull(Me.Route.Value) Or IsNull(Me.RequiredServices.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    Else
        ' Add a New record To the table
        Dim db As DAO.Database
        Dim rs As DAO.Recordset

        Set db = CurrentDb
        Set rs = db.OpenRecordset("NewFlightRequestsT", dbOpenDynaset)

        rs.AddNew

        ' Assign values To the fields in the table
        With rs
            .Fields("RefNo").Value = Me.RefNo.Value
            .Fields("ReqNo").Value = Me.ReqNo.Value
            .Fields("ReqDate").Value = Me.ReqDate.Value
            .Fields("OPSDate").Value = Me.OPSDate.Value
            .Fields("Customer").Value = Me.Customer.Value
            .Fields("Operator").Value = Me.Operator.Value
            .Fields("ACType").Value = Me.ACType.Value
            .Fields("ACMTOW").Value = Me.ACMTOW.Value
            .Fields("ACReg").Value = Me.ACRegCb.Value
            .Fields("PFL").Value = Me.PFL.Value
            .Fields("Schedule").Value = Me.Schedule.Value
            .Fields("FuelRequest").Value = Me.FuelRequestCb.Value
            .Fields("FlightSupport").Value = Me.FlightSupportCb.Value
            .Fields("RequestStatus").Value = Me.RequestStatus.Value
            .Fields("Route").Value = Me.Route.Value
            .Fields("RequiredServices").Value = Me.RequiredServices.Value

            ' Save the New record
            .Update
        End With

        ' Clean up

        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Reset form values
        'DoCmd.RunCommand acCmdRecordsGoToNew
        'DoCmd.GoToRecord , , acNewRec
        Me.Undo
        ' Display a message indicating successful addition
        MsgBox "Data added successfully."
    End If
    Me.Refresh

End Sub

Private Sub RefreshBtn_Click()
    Me.Requery
    Me.Refresh
End Sub

Private Sub ExtBriefBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    If Not IsNull(Me.Customer) Then

        strWhere = strWhere & "([Customer] Like ""*" & Me.Customer & "*"") And "
    End If

    If Not IsNull(Me.Operator) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.Operator & "*"") And "
    End If
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
        DoCmd.OpenReport "FlightSupport_ExternalReport", acViewPreview, , strWhere
    End If
End Sub

Private Sub FlightBriefBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    If Not IsNull(Me.Customer) Then
        strWhere = strWhere & "([Customer] Like ""*" & Me.Customer & "*"") And "
    End If

    If Not IsNull(Me.Operator) Then
        strWhere = strWhere & "([Operator] Like ""*" & Me.Operator & "*"") And "
    End If
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
        DoCmd.OpenReport "FlightSupport_InternalReport", acViewPreview, , strWhere
    End If
End Sub

Private Sub Form_Load()
    LoadLastRequestBtn_Click
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

Private Sub OPSDate_Click()
    InputDateField OPSDate, "Select a date To use this on your form"
End Sub

Private Sub RefNo2_AfterUpdate()
    Dim selectedRefNo As Variant
    Me.RecordSource = "Select * FROM NewFlightRequestsT WHERE RefNo = '" & Me.RefNo2 & "';"

    selectedRefNo = Me.RefNo2.Value
    If Not IsNull(selectedRefNo) Then
        Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("RefNo").Value = selectedRefNo
            Me.Refresh
            Me.Requery
            RefNo_reselect
        Else
         Exit Sub
        End If

End Sub

Private Sub RefNo_reselect()
    Dim selectedRefNo As Variant
    Dim query As String
    'Dim selectedSect_ID As Integer
    selectedRefNo = Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("RefNo").Value ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select Sect_ID, SectorNo FROM NewSectorsT WHERE RefNo = '" & selectedRefNo & "';"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query
    'selectedSect_ID = Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("Sect_ID").Value
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("AddLocationF").Controls("Sect_ID").RowSource = query
        Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("Sect_ID").RowSource = query

            ' Requery the other subform's combo box To reflect the updated options
            Me.Refresh
End Sub

Public Sub ReqDate_Click()
    InputDateField ReqDate, "Select a date To use this on your form"
End Sub

Public Sub ResetBtn_Click()
    On Error Goto ErrorHandler

        Me.Undo
        DoCmd.CancelEvent
        DoCmd.GoToRecord , , acNewRec

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub

Private Sub LoadLastRequestBtn_Click()
    Dim MaxID As Long
    MaxID = DMax("ID", "NewFlightRequestsT")
    Me.RecordSource = "Select * FROM NewFlightRequestsT WHERE ID = " & MaxID & ";"
    ' Forms!FlightSupport_LASTFORM!FlightSupport_NewAddSectorF!RefNo = Me.RefNo
End Sub

//////////////////////////////////////////////////////////////////////

'old code
Option Compare Database
Private Sub ACRegCb_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType.Value = rs("ACType").Value
            Me.ACMTOW.Value = rs("ACMTOW").Value
            Me.Operator.Value = rs("Operator").Value
            Me.Customer.Value = rs("Customer").Value
        Else
            Me.ACType.Value = Null
            Me.ACMTOW.Value = Null
            Me.Operator.Value = Null
            Me.Customer.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ACType.Value = Null
        Me.ACMTOW.Value = Null
        Me.Operator.Value = Null
        Me.Customer.Value = Null
    End If

 Exit Sub

End Sub
Private Sub cmdGenerate_Click()
    Dim rst As DAO.Recordset
    Dim randomRefValue As Integer
    Dim RefNo As String
    Dim isUnique As Boolean

    Do
        ' Generate a random number between 1 And 999
        randomRefValue = Int((9999 - 1 + 1) * Rnd + 1)

        ' Check If the generated number is greater than Or equal To 176
        If randomRefValue >= 176 Then
            ' Generate the reference number
            RefNo = "TAS-" & Format(randomRefValue, "0000") & "/" & Format(date, "yy")

            ' Check If the generated reference number already exists in the NewFlightRequestsT table
            Set rst = CurrentDb.OpenRecordset("Select RefNo FROM NewFlightRequestsT WHERE RefNo = '" & RefNo & "'", dbOpenSnapshot)

            If rst.EOF Then
                ' If the generated reference number does Not exist, assign it To the form field
                Me.RefNo.Value = RefNo

                ' Set the flag To indicate that a unique reference number has been generated
                isUnique = True
            End If

            rst.Close
            Set rst = Nothing
        End If
    Loop Until isUnique

End Sub
Private Sub AddFlightBtn_Click()
    ' Check If all the required fields are filled
    If IsNull(Me.RefNo.Value) Or IsNull(Me.ReqDate.Value) Or IsNull(Me.OPSDate.Value) Or IsNull(Me.Customer.Value) Or IsNull(Me.Operator.Value) _
        Or IsNull(Me.ACType.Value) Or IsNull(Me.ACMTOW.Value) Or IsNull(Me.ACRegCb.Value) Or IsNull(Me.PFL.Value) Or IsNull(Me.Schedule.Value) Or IsNull(Me.Route.Value) Or IsNull(Me.RequiredServices.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    Else
        ' Add a New record To the table
        Dim db As DAO.Database
        Dim rs As DAO.Recordset

        Set db = CurrentDb
        Set rs = db.OpenRecordset("NewFlightRequestsT", dbOpenDynaset)

        rs.AddNew

        ' Assign values To the fields in the table
        With rs
            .Fields("RefNo").Value = Me.RefNo.Value
            .Fields("ReqNo").Value = Me.ReqNo.Value
            .Fields("ReqDate").Value = Me.ReqDate.Value
            .Fields("OPSDate").Value = Me.OPSDate.Value
            .Fields("Customer").Value = Me.Customer.Value
            .Fields("Operator").Value = Me.Operator.Value
            .Fields("ACType").Value = Me.ACType.Value
            .Fields("ACMTOW").Value = Me.ACMTOW.Value
            .Fields("ACReg").Value = Me.ACRegCb.Value
            .Fields("PFL").Value = Me.PFL.Value
            .Fields("Schedule").Value = Me.Schedule.Value
            .Fields("FuelRequest").Value = Me.FuelRequestCb.Value
            .Fields("FlightSupport").Value = Me.FlightSupportCb.Value
            .Fields("RequestStatus").Value = Me.RequestStatus.Value
            .Fields("Route").Value = Me.Route.Value
            .Fields("RequiredServices").Value = Me.RequiredServices.Value

            ' Save the New record
            .Update
        End With

        ' Clean up

        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Reset form values
        'DoCmd.RunCommand acCmdRecordsGoToNew
        'DoCmd.GoToRecord , , acNewRec
        Me.Undo
        ' Display a message indicating successful addition
        MsgBox "Data added successfully."
    End If
    Me.Refresh

End Sub

Private Sub ExtBriefBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    If Not IsNull(Me.Customer) Then

        strWhere = strWhere & "([Customer] Like ""*" & Me.Customer & "*"") And "
    End If

    If Not IsNull(Me.Operator) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.Operator & "*"") And "
    End If
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
        DoCmd.OpenReport "FlightSupport_ExternalReport", acViewPreview, , strWhere
    End If
End Sub

Private Sub FlightBriefBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    If Not IsNull(Me.Customer) Then

        strWhere = strWhere & "([Customer] Like ""*" & Me.Customer & "*"") And "
    End If

    If Not IsNull(Me.Operator) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.Operator & "*"") And "
    End If
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
        DoCmd.OpenReport "FlightSupport_InternalReport", acViewPreview, , strWhere
    End If

End Sub

Private Sub Form_Load()
    LoadLastRequestBtn_Click
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

Private Sub OPSDate_Click()
    InputDateField OPSDate, "Select a date To use this on your form"
End Sub

Private Sub RefNo2_AfterUpdate()
    Dim selectedRefNo As Variant
    Me.RecordSource = "Select * FROM NewFlightRequestsT WHERE RefNo = '" & Me.RefNo2 & "';"

    selectedRefNo = Me.RefNo2.Value

    If Not IsNull(selectedRefNo) Then
        Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("RefNo").Value = selectedRefNo
            Me.Refresh
            Me.Requery
            RefNo_reselect

        Else
         Exit Sub
        End If

End Sub

Private Sub RefNo_reselect()
    Dim selectedRefNo As Variant
    Dim query As String
    'Dim selectedSect_ID As Integer
    selectedRefNo = Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("RefNo").Value ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select Sect_ID, SectorNo FROM NewSectorsT WHERE RefNo = '" & selectedRefNo & "';"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query
    'selectedSect_ID = Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("Sect_ID").Value
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_LASTFORM").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_LASTFORM").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("AddLocationF").Controls("Sect_ID").RowSource = query
        Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("Sect_ID").RowSource = query

            ' Requery the other subform's combo box To reflect the updated options
            ' Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("Sect_ID").Requery
            ' Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("AddLocationF").Controls("Sect_ID").Requery
            ' Forms("FlightSupport_LASTFORM").Controls("FlightSupport_LASTFORM").Form.Controls("PermitBriefF").Controls("Sect_ID").Requery
            Me.Requery
            Me.Refresh
End Sub
Private Sub RefreshBtn_Click()
    Me.Requery
    Me.Refresh
End Sub

Public Sub ReqDate_Click()
    InputDateField ReqDate, "Select a date To use this on your form"
End Sub

Public Sub ResetBtn_Click()
    On Error Goto ErrorHandler

        Me.Undo
        DoCmd.CancelEvent
        DoCmd.GoToRecord , , acNewRec

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub

Private Sub LoadLastRequestBtn_Click()
    Dim MaxID As Long
    MaxID = DMax("ID", "NewFlightRequestsT")
    Me.RecordSource = "Select * FROM NewFlightRequestsT WHERE ID = " & MaxID & ";"
    ' Forms!FlightSupport_LASTFORM!FlightSupport_NewAddSectorF!RefNo = Me.RefNo
End Sub

