Option Compare Database
Private Sub addServiceBtn_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Integer

    Dim Service As String, Vendor As String, Agent As String, ServiceStatus As String, PaymentMethod As String
    Dim Sect_ID As Variant ' Change the data type To Variant

    ' Open the main table's recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("ServicesT", dbOpenDynaset)

    ' Loop through the sets of text boxes And add records
    For i = 1 To 8 ' Change 5 To the desired number of services

        ' Retrieve values from text boxes

        Service = Nz(Me.Controls("Service" & i).Value, "")
        Vendor = Nz(Me.Controls("Vendor" & i).Value, "")
        Agent = Nz(Me.Controls("Agent" & i).Value, "")
        ServiceStatus = Nz(Me.Controls("ServiceStatus" & i).Value, "")
        PaymentMethod = Nz(Me.Controls("PaymentMethod" & i).Value, "")
        Notes = Nz(Me.Controls("Notes" & i).Value, "")
        Sect_ID = Nz(Me.Controls("Sect_ID" & i).Value, "")


        ' Check If at least one field is Not empty
        If Service <> "" Or Vendor <> "" Or Agent <> "" Or ServiceStatus <> "" Or PaymentMethod <> "" Or Notes <> "" Or Sect_ID <> "" Then
            ' Add a record
            rs.AddNew
            ' Assign Sect_ID To the field in the table
            rs.Fields("Service").Value = Service
            rs.Fields("Vendor").Value = Vendor
            rs.Fields("Agent").Value = Agent
            rs.Fields("ServiceStatus").Value = ServiceStatus
            rs.Fields("PaymentMethod").Value = PaymentMethod
            rs.Fields("Notes").Value = Notes
            rs.Fields("Sect_ID").Value = Sect_ID
            rs.Update
        End If

        ' Clear the text boxes after adding a record

        Me.Controls("Service" & i).Value = Null
        Me.Controls("Vendor" & i).Value = Null
        Me.Controls("Agent" & i).Value = Null
        Me.Controls("ServiceStatus" & i).Value = Null
        Me.Controls("PaymentMethod" & i).Value = Null
        Me.Controls("Notes" & i).Value = Null
        Me.Controls("Sect_ID" & i).Value = Null
    Next i

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Optional: Display a message indicating successful addition
    MsgBox "Service records added successfully.", vbInformation
End Sub

Private Sub followUpMark1_AfterUpdate()
    UpdateToggleButtonCaption followUpMark1
End Sub

Private Sub followUpMark2_AfterUpdate()
    UpdateToggleButtonCaption followUpMark2
End Sub

Private Sub followUpMark3_AfterUpdate()
    UpdateToggleButtonCaption followUpMark3
End Sub

Private Sub followUpMark4_AfterUpdate()
    UpdateToggleButtonCaption followUpMark4
End Sub

Private Sub followUpMark5_AfterUpdate()
    UpdateToggleButtonCaption followUpMark5
End Sub

Private Sub followUpMark6_AfterUpdate()
    UpdateToggleButtonCaption followUpMark6
End Sub

Private Sub followUpMark7_AfterUpdate()
    UpdateToggleButtonCaption followUpMark7
End Sub

Private Sub followUpMark8_AfterUpdate()
    UpdateToggleButtonCaption followUpMark8
End Sub

Private Sub followUpMark9_AfterUpdate()
    UpdateToggleButtonCaption followUpMark9
End Sub

Private Sub followUpMark10_AfterUpdate()
    UpdateToggleButtonCaption followUpMark10
End Sub

Private Sub UpdateToggleButtonCaption(toggleButton As toggleButton)
    If Not IsNull(toggleButton.Value) Then
        If toggleButton.Value Then
            toggleButton.Caption = "P"
        Else
            toggleButton.Caption = ""
        End If
    End If
End Sub
Private Sub Form_Current()
    followUpMark1_AfterUpdate
    followUpMark2_AfterUpdate
    followUpMark3_AfterUpdate
    followUpMark4_AfterUpdate
    followUpMark5_AfterUpdate
    followUpMark6_AfterUpdate
    followUpMark7_AfterUpdate
    followUpMark8_AfterUpdate
    followUpMark9_AfterUpdate
    followUpMark10_AfterUpdate
End Sub

Private Sub Form_Load()
    DoCmd.GoToRecord , , acNewRec
End Sub


With the toggle button added
    ///////////////////////////////////////

Private Sub addServiceBtn_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Integer

    Dim Service As String, Vendor As String, Agent As String, ServiceStatus As String, PaymentMethod As String
    Dim Sect_ID As Variant

    Dim toggleButton As MSForms.ToggleButton
    Dim toggleButtonName As String

    ' Open the main table's recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("ServicesT", dbOpenDynaset)

    ' Loop through the sets of text boxes And add records
    For i = 1 To 8 ' Change 5 To the desired number of services

        ' Retrieve values from text boxes
        Service = Nz(Me.Controls("Service" & i).Value, "")
        Vendor = Nz(Me.Controls("Vendor" & i).Value, "")
        Agent = Nz(Me.Controls("Agent" & i).Value, "")
        ServiceStatus = Nz(Me.Controls("ServiceStatus" & i).Value, "")
        PaymentMethod = Nz(Me.Controls("PaymentMethod" & i).Value, "")
        Notes = Nz(Me.Controls("Notes" & i).Value, "")
        Sect_ID = Nz(Me.Controls("Sect_ID" & i).Value, "")

        ' Retrieve the corresponding toggle button
        toggleButtonName = "followUpMark" & i
        Set toggleButton = Me.Controls(toggleButtonName)

        ' Check If at least one field is Not empty
        If Service <> "" Or Vendor <> "" Or Agent <> "" Or ServiceStatus <> "" Or PaymentMethod <> "" Or Notes <> "" Or Sect_ID <> "" Then
            ' Add a record
            rs.AddNew
            ' Assign values To the fields in the table
            rs.Fields("Service").Value = Service
            rs.Fields("Vendor").Value = Vendor
            rs.Fields("Agent").Value = Agent
            rs.Fields("ServiceStatus").Value = ServiceStatus
            rs.Fields("PaymentMethod").Value = PaymentMethod
            rs.Fields("Notes").Value = Notes
            rs.Fields("Sect_ID").Value = Sect_ID

            ' Update the toggle button value
            If Not IsNull(toggleButton.Value) Then
                rs.Fields("FollowUp").Value = IIf(toggleButton.Value, 1, 0)
            End If

            rs.Update
        End If

        ' Clear the text boxes And toggle button after adding a record
        Me.Controls("Service" & i).Value = Null
        Me.Controls("Vendor" & i).Value = Null
        Me.Controls("Agent" & i).Value = Null
        Me.Controls("ServiceStatus" & i).Value = Null
        Me.Controls("PaymentMethod" & i).Value = Null
        Me.Controls("Notes" & i).Value = Null
        Me.Controls("Sect_ID" & i).Value = Null
        toggleButton.Value = Null
    Next i

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Optional: Display a message indicating successful addition
    MsgBox "Service records added successfully.", vbInformation
End Sub


////////////////////////

Option Compare Database
Private CancelSave As Boolean

Private Sub AddServicesBtn_Click()
    DoCmd.OpenForm "multiNewServicesF", WindowMode:=acDialog
End Sub

Private Sub closeBtn_Click()
    DoCmd.Close acForm, "AddSectorF"
End Sub
Private Sub Form_Load()
    ' Check if the form is in DataEntry mode
    If Me.NewRecord Then
        ' Get the latest RefNo value from FlightRequestsT table
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim strSQL As String
        Dim latestRefNo As Variant

        Set db = CurrentDb
        strSQL = "SELECT TOP 1 RefNo FROM FlightRequestsT ORDER BY AddedTime DESC"
        Set rs = db.OpenRecordset(strSQL)

        ' Check if there is a latest RefNo value
        If Not rs.EOF Then
            latestRefNo = rs.Fields("RefNo").Value

            ' Assign the latest RefNo value to the form control
            Me.RefNo.Value = latestRefNo
        End If

        ' Generate a new random Sect_ID value
        Me.Sect_ID.Value = GenerateNewSect_ID()

        ' Clean up
        rs.Close
        Set rs = Nothing
        Set db = Nothing

        ' Set the CancelSave flag to True
        CancelSave = True
    End If
End Sub

Private Function GenerateNewSect_ID() As String
    ' Calculate the New RefNo value based on existing records
    Dim lastSect_ID As Variant
    lastSect_ID = DMax("Sect_ID", "SectorsT")

    Dim newSect_ID As String

    If IsNull(lastSect_ID) Then
        ' If the table is empty, Set the initial reference number
        newSect_ID = "00001"
    Else
        ' Generate a random number between 1 And 999999
        Dim randomNumber As Long
        Randomize  ' Initialize the random number generator
        randomNumber = Int((99999 - 1 + 1) * Rnd + 1)

        ' Combine random number With the prefix And format it
        newSect_ID = Format(randomNumber, "00000")
    End If

    ' Return the New RefNo value
    GenerateNewSect_ID = newSect_ID
End Function
Private Sub addSectorBrn_Click()
    ' Check if the CancelSave flag is set to True
    If CancelSave = True Then
        ' Set the CancelSave flag back to False
        CancelSave = False
        Exit Sub ' Exit the event without adding the record
    End If

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' Open the main table's recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SectorsT", dbOpenDynaset)

    ' Add a New record
    rs.AddNew

    ' Assign values To the fields in the table
    rs.Fields("Sect_ID").Value = Nz(Me.Sect_ID.Value, "")
    rs.Fields("RefNo").Value = Nz(Me.RefNo.Value, "")
    rs.Fields("SectorNo").Value = Nz(Me.SectorNo.Value, "")
    rs.Fields("SectorLocation").Value = Nz(Me.SectorLocation.Value, "")
    rs.Fields("Location").Value = Nz(Me.Location.Value, "")
    rs.Fields("ETD").Value = Nz(Me.ETD.Value, "")
    rs.Fields("ETA").Value = Nz(Me.ETA.Value, "")
    rs.Fields("ATD").Value = Nz(Me.ATD.Value, "")
    rs.Fields("ATA").Value = Nz(Me.ATA.Value, "")

    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset the form fields
    Me.SectorNo.Value = Null
    Me.SectorLocation.Value = Null
    Me.Location.Value = Null
    Me.ETD.Value = Null
    Me.ETA.Value = Null
    Me.ATD.Value = Null
    Me.ATA.Value = Null

    ' Optional: Display a message indicating successful addition
    MsgBox "Sector added successfully.", vbInformation
End Sub


