Private Sub PreformaReportBtn_Click()
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
        strWhere = strWhere & "([FuelReleaseT_New_V12].[DepartDate] >= #" & Format(Me.StartDate, "yyyy\/mm\/dd") & "#) And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([FuelReleaseT_New_V12].[DepartDate] < #" & Format(Me.EndDate + 1, "yyyy\/mm\/dd") & "#) And "
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
        qdf.sql = "Select FlightSupport_FlightRequestsT.RefNo, FlightSupport_FlightRequestsT.ReqNo, " & _
        "FlightSupport_SectorsQ.CustomerName, FuelSupport_FuelQuotationsT.Customer, FlightSupport_SectorsQ.ACReg, " & _
        "FlightSupport_SectorsQ.Operator, FuelBriefT_New.Vendor As ActualVendor, " & _
        "IIf(FuelBriefT_New.VendorPrice > FuelBriefT_New.EstimatedTotal, FuelBriefT_New.VendorPrice, FuelBriefT_New.EstimatedTotal) As ActualVendorPrice, " & _
        "FuelBriefT_New.AddedRate As ActualAddedRate, FuelBriefT_New.TASPrice As ActualTASPrice, FuelBriefT_New.FuelQuote_RefNo, " & _
        "FuelSupport_FuelQuotationsT.Quote_RefNo, FuelSupport_FuelQuotationsT.Vendor As QuotedVendor, " & _
        "IIf(FuelSupport_FuelQuotationsT.VendorPrice > FuelSupport_FuelQuotationsT.EstimatedTotal, FuelSupport_FuelQuotationsT.VendorPrice, FuelSupport_FuelQuotationsT.EstimatedTotal) As QuotedVendorPrice, " & _
        "FuelSupport_FuelQuotationsT.AddedRate As QuotedAddedRate, FuelSupport_FuelQuotationsT.Q_TASPrice, " & _
        "FuelSupport_FuelQuotationsT.Validity As Q_Validity, FuelSupport_FuelQuotationsT.Currency As Q_Currency, " & _
        "FuelBriefT_New.Location, FuelReleaseT_New_V12.DepartDate, FuelReleaseT_New_V12.NextDest, " & _
        "FuelBriefT_New.EstimatedQuantity As Estimate_ActualQuantity, FuelBriefT_New.EstimatedCost As Estimate_ActualCost, " & _
        "FuelBriefT_New.FuelStatus, FuelBriefT_New.FuelTicketStatus " & _
        "FROM FuelSupport_FuelQuotationsT " & _
        "INNER JOIN (((FlightSupport_FlightRequestsT " & _
        "INNER JOIN FlightSupport_SectorsQ ON FlightSupport_FlightRequestsT.RefNo = FlightSupport_SectorsQ.FlightRefNo) " & _
        "INNER JOIN FuelBriefT_New ON FlightSupport_SectorsQ.Origin_Location_ID = FuelBriefT_New.Location_ID) " & _
        "INNER JOIN FuelReleaseT_New_V12 ON FuelBriefT_New.Location_ID = FuelReleaseT_New_V12.Location_ID) " & _
        "ON FuelSupport_FuelQuotationsT.Quote_RefNo = FuelBriefT_New.FuelQuote_RefNo " & _
        "And FuelSupport_FuelQuotationsT.Airport = FuelBriefT_New.Location " & _
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


qdf.sql = "Select FlightSupport_FlightRequestsT.RefNo, FlightSupport_FlightRequestsT.ReqNo, " & _
"FlightSupport_SectorsQ.CustomerName, FuelSupport_FuelQuotationsT.Customer, " & _
"FlightSupport_SectorsQ.ACReg, FlightSupport_SectorsQ.Operator, FuelBriefT_New.Vendor As ActualVendor, " & _
"IIf(FuelBriefT_New.VendorPrice > FuelBriefT_New.EstimatedTotal, FuelBriefT_New.VendorPrice, FuelBriefT_New.EstimatedTotal) As ActualVendorPrice, " & _
"FuelBriefT_New.AddedRate As ActualAddedRate, FuelBriefT_New.TASPrice As ActualTASPrice, FuelBriefT_New.FuelQuote_RefNo, " & _
"FuelSupport_FuelQuotationsT.Quote_RefNo, FuelSupport_FuelQuotationsT.Vendor As QuotedVendor, " & _
"IIf(FuelSupport_FuelQuotationsT.VendorPrice > FuelSupport_FuelQuotationsT.EstimatedTotal, FuelSupport_FuelQuotationsT.VendorPrice, FuelSupport_FuelQuotationsT.EstimatedTotal) As QuotedVendorPrice, " & _
"FuelSupport_FuelQuotationsT.AddedRate As QuotedAddedRate, FuelSupport_FuelQuotationsT.Q_TASPrice, " & _
"FuelSupport_FuelQuotationsT.Validity As Q_Validity, FuelSupport_FuelQuotationsT.Currency As Q_Currency, " & _
"FuelBriefT_New.Location, FuelReleaseT_New_V12.DepartDate, FuelReleaseT_New_V12.NextDest, " & _
"FuelBriefT_New.EstimatedQuantity As Estimate_ActualQuantity, FuelBriefT_New.EstimatedCost As Estimate_ActualCost, " & _
"FuelBriefT_New.FuelStatus, FuelBriefT_New.FuelTicketStatus " & _
"FROM (((FlightSupport_FlightRequestsT " & _
"INNER JOIN FlightSupport_SectorsQ ON FlightSupport_FlightRequestsT.RefNo = FlightSupport_SectorsQ.FlightRefNo) " & _
"INNER JOIN FuelBriefT_New ON FlightSupport_SectorsQ.Origin_Location_ID = FuelBriefT_New.Location_ID) " & _
"INNER JOIN FuelReleaseT_New_V12 ON FuelBriefT_New.Location_ID = FuelReleaseT_New_V12.Location_ID) " & _
"LEFT JOIN FuelSupport_FuelQuotationsT ON FuelSupport_FuelQuotationsT.Airport = FuelBriefT_New.Location " & _
"WHERE " & strWhere & " ORDER BY FuelReleaseT_New_V12.DepartDate ASC;"


]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

Private Sub TripStatuscmb_AfterUpdate()
    Dim newStatus As String
    Dim sectID As Long
    Dim selectedStatus As String

    selectedStatus = Me.TripStatuscmb.Value

    If Not IsNull(selectedStatus) Then
        ' Check For invalid sectID once
        sectID = Nz(Me.Sect_ID, 0)
        If sectID = 0 Then
            MsgBox "Error: Sect_ID is invalid.", vbCritical
         Exit Sub
        End If

        ' Confirmation message
        If MsgBox("Are you sure you want To proceed With the update?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
            Select Case selectedStatus
             Case "CANCELED"
                newStatus = selectedStatus
             Case "CONFIRMED"
                newStatus = selectedStatus
             Case "IN PROGRESS"
                newStatus = selectedStatus
             Case Else
                ' Handle other potential values of selectedStatus
             Exit Sub
            End Select

            UpdatePermitServices newStatus, sectID
            UpdateLocationServices newStatus, sectID
            Me.Requery
        End If
    Else
     Exit Sub
    End If

 Exit Sub


End Sub

Private Sub UpdatePermitServices(newStatus As String, sectID As Long)
    Dim strSQL As String
    Dim db As DAO.Database

    Set db = CurrentDb
    strSQL = "UPDATE PermitBriefT_New Set PermitStatus = '" & newStatus & "' WHERE Sect_ID = " & sectID
    db.Execute strSQL, dbFailOnError

    Set db = Nothing
 Exit Sub

End Sub

Private Sub UpdateLocationServices(newStatus As String, sectID As Long)
    Dim strSQL As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim locID As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset("Select Loc_ID FROM FlightSupport_AlllocationsQ_V12 WHERE Sect_ID = " & sectID)

    Do While Not rs.EOF
        locID = rs!Loc_ID
        If Not IsNull(locID) And IsNumeric(locID) Then
            strSQL = "UPDATE HandlingT_New Set ServiceStatus = '" & newStatus & "' WHERE Location_ID = " & locID
            db.Execute strSQL, dbFailOnError

            strSQL = "UPDATE FuelBriefT_New Set FuelStatus = '" & newStatus & "' WHERE Location_ID = " & locID
            db.Execute strSQL, dbFailOnError

            strSQL = "UPDATE ConciergeBriefT_New Set ServiceStatus = '" & newStatus & "' WHERE Location_ID = " & locID
            db.Execute strSQL, dbFailOnError
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
 Exit Sub

 ErrorHandler:
    MsgBox "An error occurred in UpdateLocationServices: " & Err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then rs.Close
        Set rs = Nothing
        Set db = Nothing
End Sub


/////////////////////////// FUELVENDORST QUERY
Private Sub Command0_Click()
    GenerateDynamicFuelQuery
End Sub

Sub GenerateDynamicFuelQuery()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim strJoins As String
    Dim strSelect As String
    Dim strWhere As String
    Dim strOrderBy As String
    Dim fuelTable As String

    ' Initialize the database And recordset
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("Select FuelTablesList FROM FuelVendorsT")

    ' Initialize the SQL parts
    strSelect = "Select DISTINCT _Test.Airport, "
    strJoins = "FROM _Test "
    strWhere = "WHERE "
    strOrderBy = "ORDER BY _Test.Airport ASC;"

    ' Loop through the list of fuel tables And construct the dynamic SQL parts
    Do While Not rs.EOF
        fuelTable = rs!FuelTablesList

        ' Add the Select part
        strSelect = strSelect & fuelTable & ".VendorPrice As " & fuelTable & "_BasePrice, " & _
        fuelTable & ".Validity As " & fuelTable & "_Validity, " & _
        fuelTable & ".EstimatedTotal As " & fuelTable & "_EstimatedTotal, " & _
        fuelTable & ".FlightType As " & fuelTable & "_FlightType, "

        ' Add the JOIN part
        strJoins = strJoins & "LEFT JOIN " & fuelTable & " ON (_Test.FlightType = " & fuelTable & ".FlightType) " & _
        "And (FuelSupport_PriceQ_Test.Airport = " & fuelTable & ".Airport) "

        ' Move To the Next record
        rs.MoveNext
    Loop

    ' Remove the trailing comma And space from the Select part
    strSelect = Left(strSelect, Len(strSelect) - 2)

    ' Construct the final SQL query
    strSQL = strSelect & " " & strJoins & " " & strWhere & " " & strOrderBy

    ' Output the generated SQL string For debugging
    Debug.Print strSQL

    ' Write the SQL string To a text file
    Dim filePath As String
    Dim fileNum As Integer
    filePath = Environ("USERPROFILE") & "\Documents\DynamicFuelQuery.sql"
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, strSQL
    Close #fileNum

    MsgBox "SQL query written To " & filePath

    ' Create Or update the query definition
    On Error Resume Next
    Set qdf = db.QueryDefs("DynamicFuelQuery")
    On Error Goto 0

        If qdf Is Nothing Then
            ' Create the query If it does Not exist
            On Error Resume Next
            Set qdf = db.CreateQueryDef("DynamicFuelQuery", strSQL)
            If Err.Number <> 0 Then
                MsgBox "Error creating query: " & Err.Description, vbCritical
             Exit Sub
            End If
            On Error Goto 0
                MsgBox "Query created."
            Else
                ' Update the query If it exists
                qdf.SQL = strSQL
                MsgBox "Query updated."
            End If

            ' Clean up
            rs.Close
            Set rs = Nothing
            Set db = Nothing
            Set qdf = Nothing

            ' Open the generated query
            DoCmd.OpenQuery "DynamicFuelQuery"

            MsgBox "Dynamic fuel query has been generated And opened successfully.", vbInformation
End Sub