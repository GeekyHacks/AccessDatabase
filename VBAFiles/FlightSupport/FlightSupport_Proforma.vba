Option Compare Database

Private Sub CloseBtn_Click()
    DoCmd.Close acForm, "FlightSupport_ReportGeneratorF_V12"
End Sub

Private Sub Command32_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If


    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhe
        Me.FilterOn = True
        DoCmd.OpenForm "Update_FlightSupport_Test", acViewPreview, , strWhere
    End If
End Sub

Private Sub Form_Load()
    Dim fw As New clFormWindow

    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
End Sub

Private Sub FuelReleaseBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([CustomerName] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If

    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If
    If Not IsNull(Me.ReleaseID) Then
        strWhere = strWhere & "([Location_ID] Like ""*" & Me.ReleaseID.Column(0) & "*"") And "

    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([DepartDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([DepartDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If
    If Not IsNull(Me.AirportCmb) Then

        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then

        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhe
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_FuelRelease_New", acViewPreview, , strWhere
    End If

End Sub

Private Sub FuelReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then
        strWhere = strWhere & "([CustomerName] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then

        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If

    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If
    If Not IsNull(Me.ServiceStatus) Then

        strWhere = strWhere & "([FuelStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([ReqDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([ReqDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If
    If Not IsNull(Me.AirportCmb) Then

        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhe
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_FuelReport", acViewPreview, , strWhere
    End If

End Sub

Private Sub GenerateINReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([CustomerName] Like ""*" & Me.CustomerName & "*"") And "
    End If

    If Not IsNull(Me.OperatorListCb) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then

        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If

    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If
    If Not IsNull(Me.ServiceStatus) Then

        strWhere = strWhere & "([RequestStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([ReqDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([ReqDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If
    If Not IsNull(Me.AirportCmb) Then

        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
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

Private Sub GenerateOUTReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([CustomerName] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ServiceStatus) Then

        strWhere = strWhere & "([RequestStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If
    If Not IsNull(Me.AirportCmb) Then

        strWhere = strWhere & "(left([Location],8) Like ""*" & Me.AirportCmb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then

        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If


    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([ReqDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([ReqDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhe
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_FullExternalReport_V12", acViewPreview, , strWhere
    End If

End Sub

Private Sub PreformaCreatedBtn_Click()
    Dim strWhere As String
    Dim InLeng As Long
    Dim strQuery As String
    Dim strFilePath As String
    Dim strFileName As String
    Dim qdf As QueryDef
    Dim tempQueryName As String

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    ' Specify the base query name
    strQuery = "FlightSupport_FuelPreformaQ_New"

    ' Build the file name With a timestamp And customer name
    strFileName = Me.CustomerName & " " & Format(Me.StartDate, "dd_mm_yyyy") & "-" & Format(Me.EndDate, "dd_mm_yyyy") & " Fuel Release Proforma " & Format(Now(), "ddmmyyyy") & ".xlsx"

    ' Specify the file path in the current user's Documents directory
    strFilePath = Environ("USERPROFILE") & "\Documents\" & strFileName

    ' Build the WHERE clause
    If Not IsNull(Me.CustomerName) Then
        strWhere = strWhere & "([Customer] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then
        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then
        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If
    'If Not IsNull(Me.RefNoList) Then
    ' strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    ' End If

    If Not IsNull(Me.RefNoList) Then
        strWhere = strWhere & "(NewFlightRequestsT.RefNo Like '*" & Me.RefNoList & "*') And "
    End If
    If Not IsNull(Me.ServiceStatus) Then
        strWhere = strWhere & "([FuelStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([FuelReleaseT_New].[DepartDate] >= #" & Format(Me.StartDate, "yyyy\/mm\/dd") & "#) And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([FuelReleaseT_New].[DepartDate] < #" & Format(Me.EndDate + 1, "yyyy\/mm\/dd") & "#) And "
    End If
    If Not IsNull(Me.AirportCmb) Then
        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If
    If Not IsNull(Me.AirportCmb) Then
        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If

    ' Trim the trailing " And "
    InLeng = Len(strWhere) - 5
    If InLeng <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        'CurrentDb.QueryDefs.Delete tempQueryName
        strWhere = Left$(strWhere, InLeng)
        tempQueryName = "Temp_" & strQuery

        ' Check If the query already exists
        If DCount("*", "MSysObjects", "([Name] = '" & tempQueryName & "' And [Type] = 5)") > 0 Then
            ' If it exists, delete it
            CurrentDb.QueryDefs.Delete tempQueryName
        End If
        'RefNo   ReqNo   Customer    Operator    ACType  Vendor  AddedRate   TASPrice    VendorPrice EstimatedTotal  Validity    Currency    ActualQuantity  EstimatedCost   DepartDate  Airport ACReg   NextDest
        ' Now you can safely create your query
        Set qdf = CurrentDb.CreateQueryDef(tempQueryName)
        qdf.sql = "Select FlightSupport_FlightRequestsT.RefNo, FuelReleaseT_New.[Created By],FuelReleaseT_New.[Modified By], FuelReleaseT_New.Created, FlightSupport_FlightRequestsT.ReqNo, FlightSupport_FlightRequestsT.OPSDate, FlightSupport_SectorsT.Customer, FlightSupport_SectorsT.Operator, FlightSupport_SectorsT.ACType, FlightSupport_SectorsT.ACReg, " & _
        "FuelBriefT_New.Vendor, FuelBriefT_New.AddedRate, FuelBriefT_New.TASPrice, FuelBriefT_New.VendorPrice, FuelBriefT_New.EstimatedTotal, FuelBriefT_New.Validity, FuelBriefT_New.Currency, FuelBriefT_New.EstimatedQuantity, " & _
        "FuelBriefT_New.EstimatedCost, FuelReleaseT_New.DepartDate, FuelBriefT_New.Location,FuelReleaseT_New.NextDest, FuelBriefT_New.FuelStatus " & _
        "FROM ((FlightSupport_FlightRequestsT INNER JOIN FlightSupport_SectorsT ON FlightSupport_FlightRequestsT.RefNo = FlightSupport_SectorsT.RequestRefNo) INNER JOIN LocationsT_New ON FlightSupport_SectorsT.Sect_ID = LocationsT_New.Sect_ID) " & _
        "INNER JOIN (FuelBriefT_New INNER JOIN FuelReleaseT_New ON FuelBriefT_New.Location_ID = FuelReleaseT_New.Location_ID) ON LocationsT_New.ID = FuelBriefT_New.Location_ID " & _
        "WHERE FlightSupport_FlightRequestsT.RefNo Like '*" & Me.RefNoList & "*' " & _
        "And " & strWhere & _
        "ORDER BY FuelReleaseT_New.DepartDate ASC"

        ' Export the temporary query To Excel
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tempQueryName, strFilePath, True

        ' Delete the temporary query
        CurrentDb.QueryDefs.Delete tempQueryName

        ' Open the folder directory
        Application.FollowHyperlink Environ("USERPROFILE") & "\Documents\", NewWindow:=True

        Set qdf = Nothing
    End If
End Sub

Private Sub PreformaModifiedBtn_Click()
    Dim strWhere As String
    Dim InLeng As Long
    Dim strQuery As String
    Dim strFilePath As String
    Dim strFileName As String
    Dim qdf As QueryDef
    Dim tempQueryName As String

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    ' Specify the base query name
    strQuery = "FlightSupport_FuelPreformaQ"

    ' Build the file name With a timestamp And customer name
    strFileName = Me.CustomerName & " " & Format(Me.StartDate, "dd_mm_yyyy") & "-" & Format(Me.EndDate, "dd_mm_yyyy") & " Fuel Release Proforma " & Format(Now(), "ddmmyyyy") & ".xlsx"

    ' Specify the file path in the current user's Documents directory
    strFilePath = Environ("USERPROFILE") & "\Documents\" & strFileName

    ' Build the WHERE clause
    If Not IsNull(Me.CustomerName) Then
        strWhere = strWhere & "([Customer] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then
        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then
        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If
    'If Not IsNull(Me.RefNoList) Then
    ' strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    ' End If

    If Not IsNull(Me.RefNoList) Then
        strWhere = strWhere & "(FlightSupport_FlightRequestsT.RefNo Like '*" & Me.RefNoList & "*') And "
    End If
    If Not IsNull(Me.ServiceStatus) Then
        strWhere = strWhere & "([FuelStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([FuelReleaseT_New].[DepartDate] >= #" & Format(Me.StartDate, "yyyy\/mm\/dd") & "#) And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([FuelReleaseT_New].[DepartDate] < #" & Format(Me.EndDate + 1, "yyyy\/mm\/dd") & "#) And "
    End If
    If Not IsNull(Me.AirportCmb) Then
        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If
    If Not IsNull(Me.AirportCmb) Then
        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If

    ' Trim the trailing " And "
    InLeng = Len(strWhere) - 5
    If InLeng <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        'CurrentDb.QueryDefs.Delete tempQueryName
        strWhere = Left$(strWhere, InLeng)
        tempQueryName = "Temp_" & strQuery

        ' Check If the query already exists
        If DCount("*", "MSysObjects", "([Name] = '" & tempQueryName & "' And [Type] = 5)") > 0 Then
            ' If it exists, delete it
            CurrentDb.QueryDefs.Delete tempQueryName
        End If
        'RefNo   ReqNo   Customer    Operator    ACType  Vendor  AddedRate   TASPrice    VendorPrice EstimatedTotal  Validity    Currency    ActualQuantity  EstimatedCost   DepartDate  Airport ACReg   NextDest
        ' Now you can safely create your query
        Set qdf = CurrentDb.CreateQueryDef(tempQueryName)
        qdf.sql = "Select FlightSupport_FlightRequestsT.RefNo, FuelReleaseT_New.[Modified By],FuelReleaseT_New.Modified, FuelReleaseT_New.[Created By], FlightSupport_FlightRequestsT.ReqNo, FlightSupport_FlightRequestsT.OPSDate,  FlightSupport_SectorsT.Customer, FlightSupport_SectorsT.Operator, FlightSupport_SectorsT.ACType, FlightSupport_SectorsT.ACReg, " & _
        "FuelBriefT_New.Vendor, FuelBriefT_New.AddedRate, FuelBriefT_New.TASPrice, FuelBriefT_New.VendorPrice, FuelBriefT_New.EstimatedTotal, FuelBriefT_New.Validity, FuelBriefT_New.Currency, FuelBriefT_New.EstimatedQuantity, " & _
        "FuelBriefT_New.EstimatedCost, FuelReleaseT_New.DepartDate, FuelBriefT_New.Location,FuelReleaseT_New.NextDest, FuelBriefT_New.FuelStatus " & _
        "FROM ((FlightSupport_FlightRequestsT INNER JOIN FlightSupport_SectorsT ON FlightSupport_FlightRequestsT.RefNo = FlightSupport_SectorsT.RequestRefNo) INNER JOIN LocationsT_New ON FlightSupport_SectorsT.Sect_ID = LocationsT_New.Sect_ID) " & _
        "INNER JOIN (FuelBriefT_New INNER JOIN FuelReleaseT_New ON FuelBriefT_New.Location_ID = FuelReleaseT_New.Location_ID) ON LocationsT_New.ID = FuelBriefT_New.Location_ID " & _
        "WHERE FlightSupport_FlightRequestsT.RefNo Like '*" & Me.RefNoList & "*' " & _
        "And " & strWhere & _
        "ORDER BY FuelReleaseT_New.DepartDate ASC"

        ' Export the temporary query To Excel
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tempQueryName, strFilePath, True

        ' Delete the temporary query
        CurrentDb.QueryDefs.Delete tempQueryName

        ' Open the folder directory
        Application.FollowHyperlink Environ("USERPROFILE") & "\Documents\", NewWindow:=True

        Set qdf = Nothing
    End If

End Sub
Private Sub PreformaReportBtn_Click()
    Dim strWhere As String
    Dim InLeng As Long
    Dim strQuery As String
    Dim strFilePath As String
    Dim strFileName As String
    Dim qdf As QueryDef
    Dim tempQueryName As String

    Const conJetDate = "\#dd\-mmm\-yyyy\#"

    ' Specify the base query name
    strQuery = "FlightSupport_FuelPreformaQ_New"

    ' Build the file name With a timestamp And customer name
    strFileName = Me.CustomerName & " " & Format(Me.StartDate, "dd_mm_yyyy") & "-" & Format(Me.EndDate, "dd_mm_yyyy") & " Fuel Release Proforma " & Format(Now(), "ddmmyyyy") & ".xlsx"

    ' Specify the file path in the current user's Documents directory
    strFilePath = Environ("USERPROFILE") & "\Documents\" & strFileName

    ' Build the WHERE clause
    If Not IsNull(Me.CustomerName) Then
        strWhere = strWhere & "([FlightSupport_FlightRequestsT].[CustomerName] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then
        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then
        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If
    If Not IsNull(Me.RefNoList) Then
        strWhere = strWhere & "(FlightSupport_FlightRequestsT.RefNo Like '*" & Me.RefNoList & "*') And "
    End If
    If Not IsNull(Me.ServiceStatus) Then
        strWhere = strWhere & "([FuelStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([FuelReleaseT_New_V12].[DepartDate] >= #" & Format(Me.StartDate, "dd\-mmm\-yyyy") & "#) And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([FuelReleaseT_New_V12].[DepartDate] < #" & Format(Me.EndDate + 1, "dd\-mmm\-yyyy") & "#) And "
    End If
    If Not IsNull(Me.AirportCmb) Then
        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If

    ' Trim the trailing " And "
    InLeng = Len(strWhere) - 5
    If InLeng <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, InLeng)
        tempQueryName = "Temp_" & strQuery

        ' Check If the query already exists
        If DCount("*", "MSysObjects", "([Name] = '" & tempQueryName & "' And [Type] = 5)") > 0 Then
            ' If it exists, delete it
            CurrentDb.QueryDefs.Delete tempQueryName
        End If

        ' Now you can safely create your query
        Set qdf = CurrentDb.CreateQueryDef(tempQueryName)
        qdf.sql = "Select DISTINCT FlightSupport_FlightRequestsT.RefNo, FlightSupport_FlightRequestsT.ReqNo, " & _
        "FlightSupport_SectorsQ.CustomerName, FuelSupport_FuelQuotationsT.Customer, " & _
        "FlightSupport_SectorsQ.ACReg, FlightSupport_SectorsQ.Operator,  " & _
        "FuelSupport_FuelQuotationsT.Quote_RefNo, FuelSupport_FuelQuotationsT.Vendor As QuotedVendor, FuelBriefT_New.Vendor As ActualVendor," & _
        "IIf(FuelBriefT_New.VendorPrice > FuelBriefT_New.EstimatedTotal, FuelBriefT_New.VendorPrice, FuelBriefT_New.EstimatedTotal) As ActualVendorPrice, " & _
        "( FuelSupport_FuelQuotationsT.Q_TASPrice - IIf(FuelBriefT_New.VendorPrice > FuelBriefT_New.EstimatedTotal, FuelBriefT_New.VendorPrice, FuelBriefT_New.EstimatedTotal) ) As AddedMargin " & _
        "FuelSupport_FuelQuotationsT.Q_TASPrice, " & _
        "FuelBriefT_New.FuelQuote_RefNo, " & _
        "FuelSupport_FuelQuotationsT.Validity As Q_Validity, FuelSupport_FuelQuotationsT.Currency As Q_Currency, " & _
        "FuelBriefT_New.Location, FuelReleaseT_New_V12.DepartDate, FuelReleaseT_New_V12.NextDest, " & _
        "FuelBriefT_New.EstimatedQuantity As Estimate_ActualQuantity, FuelBriefT_New.EstimatedCost As Estimate_ActualCost, " & _
        "FuelBriefT_New.FuelStatus, FuelBriefT_New.FuelTicketStatus, " & _
        "FROM (((FlightSupport_FlightRequestsT " & _
        "INNER JOIN FlightSupport_SectorsQ ON FlightSupport_FlightRequestsT.RefNo = FlightSupport_SectorsQ.FlightRefNo) " & _
        "INNER JOIN FuelBriefT_New ON FlightSupport_SectorsQ.Origin_Location_ID = FuelBriefT_New.Location_ID) " & _
        "INNER JOIN FuelReleaseT_New_V12 ON FuelBriefT_New.Location_ID = FuelReleaseT_New_V12.Location_ID) " & _
        "LEFT JOIN FuelSupport_FuelQuotationsT ON (FuelSupport_FuelQuotationsT.Airport = FuelBriefT_New.Location " & _
        "And (FuelBriefT_New.FuelQuote_RefNo = FuelSupport_FuelQuotationsT.Quote_RefNo)) " & _
        "WHERE " & strWhere & " ORDER BY FuelReleaseT_New_V12.DepartDate ASC;"
        ' Export the temporary query To Excel
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tempQueryName, strFilePath, True

        ' Delete the temporary query
        CurrentDb.QueryDefs.Delete tempQueryName

        ' Open the folder directory
        Application.FollowHyperlink Environ("USERPROFILE") & "\Documents\", NewWindow:=True

        Set qdf = Nothing
    End If
End Sub
Private Sub RefNoList_AfterUpdate()
    Dim SelectedRefNo As Variant
    Dim query As String

    SelectedRefNo = Me.RefNoList.Value ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select Location_ID, Location FROM FuelSupport_ReleaseQ_New WHERE RefNo = '" & SelectedRefNo & "';"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query

    'Forms("LASTFORM").Controls("LastForm").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    'Forms("FlightSupport_ReportGeneratorF_New").Controls("ReleaseID").RowSource = query

    Me.ReleaseID.RowSource = query

    ' Requery the other subform's combo box To reflect the updated options
    Me.ReleaseID.Requery

End Sub


Private Sub Reports_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then
        strWhere = strWhere & "([CustomerName] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then
        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If

    If Not IsNull(Me.RefNoList) Then
        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If
    If Not IsNull(Me.ServiceStatus) Then
        strWhere = strWhere & "([FuelStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([ReqDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([ReqDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If
    If Not IsNull(Me.AirportCmb) Then
        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If
    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhe
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_ListFuelReportByDate", acViewPreview, , strWhere
    End If


End Sub

Private Sub TrackFuelReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([CustomerName] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.OperatorListCb) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ACRegCmb) Then

        strWhere = strWhere & "([ACReg] Like ""*" & Me.ACRegCmb & "*"") And "
    End If
    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If
    If Not IsNull(Me.ServiceStatus) Then

        strWhere = strWhere & "([FuelStatus] Like ""*" & Me.ServiceStatus & "*"") And "
    End If
    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([DepartDate] >= #" & Format(Me.StartDate, "yyyy\/mm\/dd") & "#) And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([DepartDate] < #" & Format(Me.EndDate + 1, "yyyy\/mm\/dd") & "#) And "
    End If
    If Not IsNull(Me.AirportCmb) Then
        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If
    If Not IsNull(Me.AirportCmb) Then

        strWhere = strWhere & "([Location] Like ""*" & Me.AirportCmb & "*"") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhe
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_TrackFuelReport", acViewPreview, , strWhere
    End If

End Sub
