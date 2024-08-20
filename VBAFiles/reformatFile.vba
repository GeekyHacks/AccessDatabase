Private Sub GenerateOUTReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.PermitNo) Then
        strWhere = strWhere & "([PermitNo] Like ""*" & Me.PermitNo & "*"") And "
    End If

    If Not IsNull(Me.CustomerName) Then
        strWhere = strWhere & "([Customer] Like ""*" & Me.CustomerName & "*"") And "

    End If

    If Not IsNull(Me.OperatorListCb) Then
        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If

    If Not IsNull(Me.ReqNoList) Then
        strWhere = strWhere & "([RequestNo] Like ""*" & Me.ReqNoList & "*"") And "
    End If

    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([OPSDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If

    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([OPSDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "OVRF_Amends_REPORT", acViewPreview, , strWhere
    End If

End Sub




Private Sub Report_Load()
    If IsNull(Me.txtAmends) Or Me.txtAmends = "" Or IsNull(Me.txtAmendsNo) Or Me.txtAmendsNo = "" Then
        Me.txtAmends.Visible = False
        Me.txtAmendsNo.Visible = False
        Me.txtCallSign.Visible = False
        Me.txtID.Visible = False
        Me.txtOPSDate.Visible = False
        Me.txtPermitNo.Visible = False
        Me.txtUpdateDate.Visible = False
        Me.txtUser.Visible = False
    End If


End Sub






On Error Resume Next
DoCmd.OpenReport "OVRF_Amends_REPORT", acViewPreview, , , acWindowNormal, "OVRF_Q"

If Err.Number <> 0 Then
    MsgBox "An error occurred While opening the report.", vbCritical, "Error"
 Exit Sub
End If

On Error Goto 0

    If CurrentProject.AllReports("OVRF_Amends_REPORT").IsLoaded Then
        Reports("OVRF_Amends_REPORT").Controls("PermitNoCount").Value = PermitCount
    End If



Private Function GetLastRefNo() As Variant

    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    'Dim lastRefNo As Variant
    Dim MaxID As Long
    MaxID = DMax("ID", "OVRF_RequestT")
    ' Me.RecordSource = "Select * FROM NewSectorsT WHERE ID = " & MaxID & ";"

    Set db = CurrentDb
    strSQL = "Select * FROM OVRF_RequestT WHERE ID = " & MaxID & ";"
    Set rs = db.OpenRecordset(strSQL)

    ' Check If a record was found
    If Not rs.EOF Then
        LastRefNo = rs.Fields("RefNo").Value
    End If

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLastRefNo = LastRefNo
End Function


Private Sub Form_current()

    Dim CAMA As Variant
    Dim LastOPS_ID As Variant
    Dim OVR_FORM As Form

    Set OVR_FORM = Forms("Update_OVR_OPSF")
    CAMA = OVR_FORM.Controls("Update_OVR_OPSF").Form.Controls("OPSPermitNo")
    LastOPS_ID = OVR_FORM.Controls("Update_OVR_OPSF").Form.Controls("OPS_ID")
End Sub

Private Sub Form_Load()
    Dim CAMA As Variant
    Dim LastOPS_ID As Variant
    Dim OVR_FORM As Form

    On Error Resume Next
    Set OVR_FORM = Forms("Update_OVR_OPSF")
    On Error Goto 0

        If OVR_FORM Is Nothing Then
            ' Form is Not loaded
            LoadLastBtn_Click
            Me.closeBtn.Visible = False
        Else
            Me.closeBtn.Visible = True
            ' Form is loaded
            If OVR_FORM.Visible = False Then
                LoadLastBtn_Click
            Else
                If OVR_FORM.Controls("Update_OVR_OPSF").Form Is Nothing Then
                    LoadLastBtn_Click
                Else
                    CAMA = OVR_FORM.Controls("Update_OVR_OPSF").Form.Controls("OPSPermitNo")
                    LastOPS_ID = OVR_FORM.Controls("Update_OVR_OPSF").Form.Controls("OPS_ID")
                    If IsNull(CAMA) Then
                        LoadLastBtn_Click
                    Else
                        Me.CAMA = CAMA
                        Me.OPS_ID = LastOPS_ID
                    End If
                End If
            End If
        End If
End Sub