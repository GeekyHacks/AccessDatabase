Option Compare Database
Option Explicit
Private Sub CustomerCmb_AfterUpdate()
    Dim Client As Variant
    Client = Me.CustomerCmb.Value
    Me.Flight_Support_PendingFuelTicketsList.Form.Filter = "CustomerName = '" & Client & "'"
    Me.Flight_Support_PendingFuelTicketsList.Form.FilterOn = True
    Me.Flight_Support_PendingFuelTicketsList.Form.Requery
End Sub
Private Sub closeBtn_Click()
    DoCmd.Close acForm, "FlightSupport_PendingFuelTickets_V12"
End Sub
Private Sub PrintBtn_Click()
    On Error GoTo ErrorHandler

    ' Specify the base query name
    Dim strQuery As String
    strQuery = "FlightSupport_PendingFuelTicketQ_V12"

    ' Initialize the WHERE clause
    Dim strWhere As String
    strWhere = ""

    ' Add filter for CustomerName if provided
    If Not IsNull(Me.CustomerCmb) Then
        strWhere = strWhere & "([FlightSupport_FlightRequestsT].[CustomerName] Like '*" & Me.CustomerCmb & "*') And "
    End If

    ' Remove the trailing " And " if any
    If Len(strWhere) > 0 Then
        strWhere = Left(strWhere, Len(strWhere) - 5)
    End If

    ' Build the final SQL query
    Dim strSQL As String
    strSQL = "SELECT FlightSupport_SectorsQ.FlightRefNo, FlightSupport_SectorsQ.ACReg, FlightSupport_SectorsQ.CallSign, FuelBriefT_New.Location, " & _
             "Format(FuelReleaseT_New_V12.DepartDate, 'yyyy-mm-dd') AS DepartDate, FuelReleaseT_New_V12.ReleaseETD, FuelBriefT_New.FuelTicketStatus, " & _
             "FlightSupport_FlightRequestsT.CustomerName, FuelBriefT_New.FuelStatus " & _
             "FROM (((FlightSupport_FlightRequestsT INNER JOIN FlightSupport_SectorsQ ON FlightSupport_FlightRequestsT.RefNo = FlightSupport_SectorsQ.FlightRefNo) " & _
             "INNER JOIN FlightSupport_AlllocationsQ_V12 ON FlightSupport_SectorsQ.Sect_ID = FlightSupport_AlllocationsQ_V12.Sect_ID) " & _
             "INNER JOIN FuelBriefT_New ON FlightSupport_AlllocationsQ_V12.Loc_ID = FuelBriefT_New.Location_ID) " & _
             "INNER JOIN FuelReleaseT_New_V12 ON FuelBriefT_New.Location_ID = FuelReleaseT_New_V12.Location_ID " & _
             "WHERE ( ((FlightSupport_SectorsQ.FlightRefNo)<>'') AND ((FuelBriefT_New.FuelStatus)<>'CANCELED') AND ((FuelBriefT_New.FuelTicketStatus)='Pending'))"

    ' Append the WHERE clause if any
    If Len(strWhere) > 0 Then
        strSQL = strSQL & " AND " & strWhere
    End If

    ' Build the file name with a timestamp and customer name
    Dim strFileName As String
    strFileName = Me.CustomerCmb & " pending fuel tickets " & Format(Now(), "ddmmyyyy") & ".xlsx"

    ' Specify the file path in the current user's Documents directory
    Dim strFilePath As String
    strFilePath = Environ("USERPROFILE") & "\Documents\" & strFileName

    ' Create a temporary query
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.CreateQueryDef("TempTicketsQuery", strSQL)

    ' Export the query results to an Excel file
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "TempTicketsQuery", strFilePath, True

    ' Delete the temporary query
    CurrentDb.QueryDefs.Delete "TempTicketsQuery"

    ' Notify the user
    MsgBox "Query results have been exported to " & strFilePath, vbInformation

    ' Open the folder directory
    Application.FollowHyperlink Environ("USERPROFILE") & "\Documents\", NewWindow:=True

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "TempTicketsQuery"
End Sub
