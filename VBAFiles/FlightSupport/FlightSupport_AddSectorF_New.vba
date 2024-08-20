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
            Me.MTOW.Value = rs("ACMTOW").Value
            Me.Operator.Value = rs("Operator").Value
            Me.Customer.Value = rs("Customer").Value
            Me.CallSigntxt.Value = rs("callSign").Value
        Else
            Me.ACType.Value = Null
            Me.MTOW.Value = Null
            Me.Operator.Value = Null
            Me.Customer.Value = Null
            Me.CallSigntxt.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ACType.Value = Null
        Me.MTOW.Value = Null
        Me.Operator.Value = Null
        Me.Customer.Value = Null
        Me.CallSigntxt.Value = Null
    End If

 Exit Sub

End Sub
Private Sub UpdateUser1()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User1 = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
Private Sub addSectID()
    Dim RefValue As Integer
    Dim rst As DAO.Recordset

    RefValue = Nz(DLookup("RefNum", "TSect_ID"), 0)

    Me.Sect_ID.Value = Format(RefValue + 1, "000")

    Set rst = CurrentDb.OpenRecordset("TSect_ID", dbOpenDynaset, dbSeeChanges)

    With rst
        .Edit
        ![RefNum] = ![RefNum] + 1
        .Update
    End With

    rst.Close
    Set rst = Nothing

End Sub
Private Sub addSectorBtn_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    addSectID

    If IsNull(Me.Sect_ID) Or IsNull(Me.SectorNo) Or IsNull(Me.CallSigntxt) Or IsNull(Me.SectorRefNo) Or IsNull(Me.ACRegCb) Or IsNull(Me.Operator) Or IsNull(Me.Customer) Or IsNull(Me.ACType) Then

        MsgBox "Please add missing fields.", vbExclamation
     Exit Sub
    End If


    ' Check If a record With the same SectorNo, Sect_ID, ORIGIN, And Destination already exists
    Set db = CurrentDb
    Set rs = db.OpenRecordset("FlightSupport_SectorsT_New", dbOpenDynaset)
    If DCount("*", "FlightSupport_SectorsT_New", "Sect_ID = " & CInt(Me.Sect_ID.Value)) > 0 Then
        MsgBox "A record With the same Sect_ID, ORIGIN, And Destination already exists.", vbExclamation, "Duplicate Record"
        rs.Close
        Set rs = Nothing
        Set db = Nothing
     Exit Sub
    End If

    Set rs = db.OpenRecordset("FlightSupport_SectorsT_New", dbOpenDynaset)
    rs.AddNew
    ' Assign values To the fields in the table

    rs.Fields("SectorStatus").Value = Me.TripStatuscmb.Value
    rs.Fields("CallSign").Value = Me.CallSigntxt.Value
    rs.Fields("Sect_ID").Value = Me.Sect_ID.Value
    rs.Fields("SectorNo").Value = Me.SectorNo.Value
    rs.Fields("FlightRefNo").Value = Me.SectorRefNo.Value
    rs.Fields("Notes").Value = Me.Notes.Value
    rs.Fields("ACReg").Value = Me.ACRegCb.Value
    rs.Fields("Operator").Value = Me.Operator.Value
    rs.Fields("ACType").Value = Me.ACType.Value
    rs.Fields("MTOW").Value = Me.MTOW.Value
    rs.Fields("Customer").Value = Me.Customer.Value

    ' Save the New record
    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    AddOrigin_Dest

    Reset_Fields

    Me.Requery
    Me.Refresh

    ' Display a message indicating successful addition
    MsgBox "Data added successfully."

End Sub

Private Sub AddOrigin_Dest()
    addOrigin
    addDest
End Sub
Private Sub addOrigin()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim boundControl As control

    If IsNull(Me.FlightSupport_AddOriginF_Form.Form.Controls("OriginTxt").Value) Or IsNull(Me.FlightSupport_AddOriginF_Form.Form.Controls("ETD").Value) Or IsNull(Me.FlightSupport_AddOriginF_Form.Form.Controls("ETA").Value) Or IsNull(Me.FlightSupport_AddOriginF_Form.Form.Controls("ArrivalDate").Value) Or IsNull(Me.FlightSupport_AddOriginF_Form.Form.Controls("DepartDate").Value) Then

        MsgBox "Please add missing fields.", vbExclamation
     Exit Sub
    End If


    ' Check If a record With the same SectorNo, Sect_ID, ORIGIN, And Destination already exists
    Set db = CurrentDb
    Set rs = db.OpenRecordset("FlightSupport_OriginT", dbOpenDynaset)


    Set boundControl = Me.FlightSupport_AddDestinationF_Form.Form.Controls("Destination")

    If DCount("*", "FlightSupport_DestinationT", "Sect_ID = " & CInt(Me.Sect_ID.Value) & " And 'OriginAirportCode' = '" & CStr(boundControl.Column(0)) & "'") > 0 Then
        MsgBox "A record With the same Sect_ID, OriginAirportCode.", vbExclamation, "Duplicate Record"
        rs.Close
        Set rs = Nothing
        Set db = Nothing
     Exit Sub
    End If

    Set rs = db.OpenRecordset("FlightSupport_OriginT", dbOpenDynaset)
    rs.AddNew
    ' Assign values To the fields in the table
    rs.Fields("OriginAirportCode").Value = Me.FlightSupport_AddOriginF_Form.Form.Controls("OriginTxt").Column(0)
    rs.Fields("AirportName").Value = Me.FlightSupport_AddOriginF_Form.Form.Controls("OriginTxt").Column(1)
    rs.Fields("Sect_ID").Value = Me.Sect_ID.Value
    rs.Fields("Origin_ETA").Value = Me.FlightSupport_AddOriginF_Form.Form.Controls("ETA").Value
    rs.Fields("Origin_ETD").Value = Me.FlightSupport_AddOriginF_Form.Form.Controls("ETD").Value
    rs.Fields("ArrivalDate").Value = Me.FlightSupport_AddOriginF_Form.Form.Controls("ArrivalDate").Value
    rs.Fields("DepartDate").Value = Me.FlightSupport_AddOriginF_Form.Form.Controls("DepartDate").Value

    ' Save the New record
    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset form values

    Me.Undo
    Me.Requery
    Me.Refresh

End Sub
Private Sub addDest()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If IsNull(Me.FlightSupport_AddDestinationF_Form.Form.Controls("Destination").Value) Or IsNull(Me.FlightSupport_AddDestinationF_Form.Form.Controls("ETA").Value) Or IsNull(Me.FlightSupport_AddDestinationF_Form.Form.Controls("ArrivalDate").Value) Then

        MsgBox "Please add missing fields.", vbExclamation
     Exit Sub
    End If

    ' Check If a record With the same SectorNo, Sect_ID, ORIGIN, And Destination already exists
    Set db = CurrentDb
    Set rs = db.OpenRecordset("FlightSupport_DestinationT", dbOpenDynaset)
    Dim boundControl As control

    Set boundControl = Me.FlightSupport_AddDestinationF_Form.Form.Controls("Destination")

    If DCount("*", "FlightSupport_DestinationT", "Sect_ID = " & CInt(Me.Sect_ID.Value) & " And 'DestAirportCode' = '" & CStr(boundControl.Column(0)) & "'") > 0 Then
        MsgBox "A record With the same Sect_ID, Destination.", vbExclamation, "Duplicate Record"
        rs.Close
        Set rs = Nothing
        Set db = Nothing
     Exit Sub
    End If

    Set rs = db.OpenRecordset("FlightSupport_DestinationT", dbOpenDynaset)
    rs.AddNew
    ' Assign values To the fields in the table
    rs.Fields("Sect_ID").Value = Me.Sect_ID.Value
    rs.Fields("DestAirportCode").Value = Me.FlightSupport_AddDestinationF_Form.Form.Controls("Destination").Column(0)
    rs.Fields("AirportName").Value = Me.FlightSupport_AddDestinationF_Form.Form.Controls("Destination").Column(1)
    rs.Fields("Destination_ETA").Value = Me.FlightSupport_AddDestinationF_Form.Form.Controls("ETA").Value
    rs.Fields("ArrivalDate").Value = Me.FlightSupport_AddDestinationF_Form.Form.Controls("ArrivalDate").Value
    ' Save the New record
    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset form values


    Me.Undo
    Me.Requery
    Me.Refresh

End Sub
Public Sub Reset_Fields()
    Me.Sect_ID.Value = Null
    Me.SectorNo.Value = Null
    Me.FlightSupport_AddDestinationF_Form.Form.Controls("Destination").Value = Null
    Me.FlightSupport_AddDestinationF_Form.Form.Controls("ETA").Value = Null
    Me.FlightSupport_AddDestinationF_Form.Form.Controls("ArrivalDate").Value = Null
    Me.FlightSupport_AddOriginF_Form.Form.Controls("ETA").Value = Null
    Me.FlightSupport_AddOriginF_Form.Form.Controls("ETD").Value = Null
    Me.FlightSupport_AddOriginF_Form.Form.Controls("ArrivalDate").Value = Null
    Me.FlightSupport_AddOriginF_Form.Form.Controls("DepartDate").Value = Null
End Sub
Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FlightSupport_AddSectorsF", acSaveNo
End Sub

Private Sub DepartDate_GotFocus()
    DepartDate_Click
End Sub

Private Sub ArrivalDate_GotFocus()
    ArrivalDate_Click
End Sub

Private Sub LoadLastBtn_Click()
    Me.RequestRefNo = GetLastRequestID
End Sub
Private Function GetLastRequestRefNo() As Variant
    Dim LastRequestID As Variant

    LastRequestID = DMax("ID", "FlightSupport_FlightRequestsT")
    GetLastRequestRefNo = LastRequestID
End Function

Public Sub DepartDate_Click()
    InputDateField DepartDate, "Select a date To use this on your form"
End Sub
Public Sub ArrivalDate_Click()
    InputDateField ArrivalDate, "Select a date To use this on your form"
End Sub

Private Sub RefreshBtn_Click()
    Me.Requery
    Me.Refresh
End Sub
Private Sub Form_Load()
    If Not IsNull(Me.RequestRefNo) Then
        Me.RequestRefNo = Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_RequestF_New").Form.Controls("RefNo").Value
        Me.SectorRefNo = Me.RequestRefNo
    Else
     Exit Sub
    End If
End Sub
Private Sub Sect_IDCmb_AfterUpdate()
    UpdateIDQueries
End Sub

Private Sub UpdateIDQueries()
    'On Error Goto ErrorHandler
    Dim selectedSect_ID As Integer

    Dim LocationsQuery As String
    Dim locSect_ID As Integer

    selectedSect_ID = Me.Sect_IDCmb.Column(1) ' Get the selected value from RefNo combo box

    LocationsQuery = "Select Location,ID FROM LocationsT_New WHERE Sect_ID = " & selectedSect_ID & ";"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID1").RowSource = LocationsQuery
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID2").RowSource = LocationsQuery
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID3").RowSource = LocationsQuery
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID4").RowSource = LocationsQuery
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID5").RowSource = LocationsQuery
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID6").RowSource = LocationsQuery

    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("FuelF").Controls("Location_ID").RowSource = LocationsQuery
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("FuelF").Controls("FlightSupport_FuelReleaseF").Controls("Location_ID").RowSource = LocationsQuery
    Forms("FlightSupport_LASTFORM_New").Controls("FlightSupport_BriefsF").Form.Controls("ConciergeBriefF").Controls("Location_ID").RowSource = LocationsQuery

    Me.Requery
    ' Assign selectedSect_ID To locSect_ID
    locSect_ID = selectedSect_ID
    ' Update the value of Sect_ID in the other subform
    Me.Controls("FlightSupport_AddLocationF").Controls("Sect_ID").Value = locSect_ID

    'ErrorHandler:
    'MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub
Private Sub selectLocations()
    Dim strSQL As String
    Dim rs As DAO.Recordset
    Dim strRowSource As String
    Dim selectedSect_ID As Integer

    ' Get the selected Sect_ID
    selectedSect_ID = Me.Sect_IDCmb.Column(1)

    ' Build the SQL query (using JOIN instead of IN)
    strSQL = "Select DISTINCT Airports_DataT.Full_Airport " & _
    "FROM Airports_DataT " & _
    "INNER JOIN FlightSupport_SectorsT ON (Airports_DataT.Full_Airport = FlightSupport_SectorsT.Origin Or Airports_DataT.Full_Airport = FlightSupport_SectorsT.Destination) " & _
    "WHERE FlightSupport_SectorsT.Sect_ID = " & selectedSect_ID & " " & _
    "ORDER BY Airports_DataT.Full_Airport;"

    ' Open the recordset
    Set rs = CurrentDb.OpenRecordset(strSQL)

    ' Build the row source string
    strRowSource = ""
    Do While Not rs.EOF
        strRowSource = strRowSource & rs!Full_Airport & ";"
        rs.MoveNext
    Loop

    ' Remove the trailing semicolon only If strRowSource is Not empty
    If Len(strRowSource) > 0 Then
        strRowSource = Left(strRowSource, Len(strRowSource) - 1)
    End If

    ' Update the combo box's row source
    Me.Controls("FlightSupport_AddLocationF").Controls("Location").RowSource = strRowSource
    Me.Controls("FlightSupport_AddLocationF").Controls("Location").ColumnCount = 1

    ' Close the recordset
    rs.Close
    Set rs = Nothing

End Sub

