
Option Compare Database

Private Sub Form_Load()
    LoadLastBtn_Click
End Sub
Private Sub LoadLastBtn_Click()
    Me.RefNo.DefaultValue = "'" & GetLastRefNo() & "'"
End Sub

Private Sub newSectorBtn_Click()
    ' Clear the form fields To prepare For a New record
    cmdGenerate_Click
End Sub

Private Sub cmdGenerate_Click()
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
Private Function GetLastRefNo() As Variant
    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Dim MaxID As Long
    MaxID = DMax("ID", "NewFlightRequestsT")

    Set db = CurrentDb
    strSQL = "Select * FROM NewFlightRequestsT WHERE ID = " & MaxID & ";"
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
Private Sub AddSectorBtn_Click()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' Me.Customer.Value = Nz(Me.CurrentARP.Value, "")
    If IsNull(Me.Sect_ID) Then
        MsgBox "Please press Add New "
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset("NewSectorsT", dbOpenDynaset)

        rs.AddNew

        ' Assign values To the fields in the table
        rs.Fields("SectorRoute").Value = Me.SectorRoute.Value
        rs.Fields("Sect_ID").Value = Me.Sect_ID.Value
        rs.Fields("SectorNo").Value = Me.SectorNo.Value
        rs.Fields("RefNo").Value = Me.RefNo.Value
        ' Save the New record
        rs.Update

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Reset form values
        Me.SectorRoute.Value = Null
        Me.Sect_ID.Value = Null
        Me.SectorNo.Value = Null
        Me.Undo
        Me.Requery
        Me.Refresh


        ' Display a message indicating successful addition
        MsgBox "Data added successfully."
        RefNo_AfterUpdate
    End If

End Sub
Private Sub RefNo_AfterUpdate()
    RefNo_reselect
End Sub
Private Sub RefNo_reselect()
    Dim selectedRefNo As Variant
    Dim query As String

    selectedRefNo = Me.RefNo.Value ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select Sect_ID, SectorNo FROM NewSectorsT WHERE RefNo = '" & selectedRefNo & "';"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query

    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Form.Controls("AddLocationF").Controls("Sect_ID").RowSource = query
        Me.Sect_ID.RowSource = query

        ' Requery the other subform's combo box To reflect the updated options
        Me.Requery
        Me.Refresh
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
Private Sub Sect_ID_AfterUpdate()
    UpdateIDQueries
End Sub

Private Sub UpdateIDQueries()
    'On Error Goto ErrorHandler
    Dim selectedSect_ID As Integer
    Dim query As String
    Dim locSect_ID As Integer

    selectedSect_ID = Me.Sect_ID.Column(0) ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select ID FROM LocationsT WHERE Sect_ID = " & selectedSect_ID & ";"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID1").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID2").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID3").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID4").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID5").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("HandlingF").Controls("Location_ID6").RowSource = query

    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("FuelF").Controls("Location_ID").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("FuelF").Controls("FlightSupport_FuelReleaseF").Controls("Location_ID").RowSource = query
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_BriefsF").Form.Controls("ConciergeBriefF").Controls("Location_ID").RowSource = query

    Me.Requery
    ' Assign selectedSect_ID To locSect_ID
    locSect_ID = selectedSect_ID

    ' Update the value of Sect_ID in the other subform
    Forms("FlightSupport_LASTFORM").Controls("FlightSupport_NewAddSectorF").Controls("AddLocationF").Controls("Sect_ID").Value = locSect_ID

        'ErrorHandler:
        'MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub



'/////////////////////////////////////////////////////////////////////
' OLD CODE


Option Compare Database

Private Sub Command32_Click()
    Me.Requery
    Me.Refresh
End Sub

Private Sub Form_Load()
    LoadLastBtn_Click
End Sub
Private Sub LoadLastBtn_Click()
    Me.RefNo.DefaultValue = "'" & GetLastRefNo() & "'"
End Sub

Private Sub newSectorBtn_Click()
    ' Clear the form fields To prepare For a New record
    cmdGenerate_Click
End Sub

Private Sub cmdGenerate_Click()
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
Private Function GetLastRefNo() As Variant
    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Dim MaxID As Long
    MaxID = DMax("ID", "NewFlightRequestsT")

    Set db = CurrentDb
    strSQL = "Select * FROM NewFlightRequestsT WHERE ID = " & MaxID & ";"
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
Private Sub AddSectorBtn_Click()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' Me.Customer.Value = Nz(Me.CurrentARP.Value, "")
    If IsNull(Me.Sect_ID) Then
        MsgBox "Please press Add New "
     Exit Sub
    Else
        Set db = CurrentDb
        Set rs = db.OpenRecordset("NewSectorsT", dbOpenDynaset)

        rs.AddNew

        ' Assign values To the fields in the table
        rs.Fields("SectorRoute").Value = Me.SectorRoute.Value
        rs.Fields("Sect_ID").Value = Me.Sect_ID.Value
        rs.Fields("SectorNo").Value = Me.SectorNo.Value
        rs.Fields("RefNo").Value = Me.RefNo.Value
        ' Save the New record
        rs.Update

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Reset form values
        Me.SectorRoute.Value = Null
        Me.Sect_ID.Value = Null
        Me.SectorNo.Value = Null
        Me.Undo
        Me.Requery
        Me.Refresh


        ' Display a message indicating successful addition
        MsgBox "Data added successfully."
        RefNo_AfterUpdate
    End If

End Sub
Private Sub RefNo_AfterUpdate()
    RefNo_reselect
End Sub
Private Sub RefNo_reselect()
    Dim selectedRefNo As Variant
    Dim query As String

    selectedRefNo = Me.RefNo.Value ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select Sect_ID, SectorNo FROM NewSectorsT WHERE RefNo = '" & selectedRefNo & "';"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query

    Forms("LASTFORM").Controls("LastForm").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("PermitBriefF").Controls("Sect_ID").RowSource = query
    Forms("LASTFORM").Controls("NewAddSectorF").Form.Controls("AddLocationF").Controls("Sect_ID").RowSource = query
        Me.Sect_ID.RowSource = query

        ' Requery the other subform's combo box To reflect the updated options
        'Me.Sect_ID.Requery
        'Forms("LASTFORM").Controls("NewAddSectorF").Form.Controls("AddLocationF").Controls("Sect_ID").Requery
        'Forms("LASTFORM").Controls("LastForm").Form.Controls("PermitBriefF").Controls("Sect_ID").Requery
        Me.Requery
        Me.Refresh
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
Private Sub Sect_ID_AfterUpdate()
    UpdateIDQueries
End Sub

Private Sub UpdateIDQueries()
    'On Error Goto ErrorHandler
    Dim selectedSect_ID As Integer
    Dim query As String
    Dim locSect_ID As Integer

    selectedSect_ID = Me.Sect_ID.Column(0) ' Get the selected value from RefNo combo box

    ' Construct the dynamic query based on the selected value
    query = "Select ID FROM LocationsT WHERE Sect_ID = " & selectedSect_ID & ";"

    ' Set the RowSource Property of the other subform's combo box To the dynamic query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID1").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID2").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID3").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID4").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID5").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID6").RowSource = query

    Forms("LASTFORM").Controls("LastForm").Form.Controls("FuelF").Controls("Location_ID").RowSource = query
    ' Forms("LASTFORM").Controls("LastForm").Form.Controls("FuelF").Form.Controls("FlightSupport_FuelReleaseF").Controls("Location_ID").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("FuelF").Controls("FlightSupport_FuelReleaseF").Controls("Location_ID").RowSource = query
    Forms("LASTFORM").Controls("LastForm").Form.Controls("ConciergeBriefF").Controls("Location_ID").RowSource = query

    ' Requery the other subform's combo box To reflect the updated options
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID1").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID2").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID3").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID4").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID5").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("HandlingF").Controls("Location_ID6").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("FuelF").Controls("Location_ID").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("FuelF").Controls("FlightSupport_FuelReleaseF").Controls("Location_ID").Requery
    Forms("LASTFORM").Controls("LastForm").Form.Controls("ConciergeBriefF").Controls("Location_ID").Requery

    ' Assign selectedSect_ID To locSect_ID
    locSect_ID = selectedSect_ID

    ' Update the value of Sect_ID in the other subform
    Forms("LASTFORM").Controls("NewAddSectorF").Controls("AddLocationF").Controls("Sect_ID").Value = locSect_ID

        'ErrorHandler:
        'MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub






