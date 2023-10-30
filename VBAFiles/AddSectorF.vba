
Option Compare Database

Private Sub AddServicesBtn_Click()
    DoCmd.OpenForm "multiNewServicesF", WindowMode:=acDialog
End Sub

Private Sub closeBtn_Click()
    ' Check If the form is in NewRecord mode before closing
    If Me.NewRecord Then
        DoCmd.Close acForm, "AddSectorF", acSaveNo
    Else
        DoCmd.Close acForm, "AddSectorF"
    End If
End Sub
Private Sub Form_Load()
    ' Navigate to a new record
    DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
    
    ' Set the RefNo value
    Me.RefNo.Value = GetLastRefNo()
    
    ' Set the Sect_ID value
    Me.Sect_ID.Value = GenerateNewSect_ID()
End Sub

Private Sub newSectorBtn_Click()
    ' Check If there are any unsaved changes
    If Me.Dirty Then
        Me.Undo ' Undo any changes To cancel the current record
    End If

    ' Clear the form fields To prepare For a New record
    ResetFormFields

    ' Generate a New Sect_ID And assign it To the form control
    Me.Sect_ID.Value = GenerateNewSect_ID()

    ' Set focus To the first field For data entry
    Me.SectorNo.SetFocus
End Sub

Private Sub addSectorBtn_Click()
 ' Rest of the code To add the record goes here...
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    ' Open the main table's recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SectorsT", dbOpenDynaset)

    ' Add a New record
    rs.AddNew

    ' Assign values To the fields in the table
    rs!Sect_ID = Me.Sect_ID.Value
    rs!RefNo = Me.RefNo.Value
    rs!SectorNo = Me.SectorNo.Value
    rs!SectorLocation = Me.SectorLocation.Value
    rs!Location = Me.Location.Value
    rs!ETD = Me.ETD.Value
    rs!ETA = Me.ETA.Value
    rs!ATD = Me.ATD.Value
    rs!ATA = Me.ATA.Value

    rs.Update

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Reset the form fields
    ResetFormFields

    ' Optional: Display a message indicating successful addition
    MsgBox "Sector added successfully.", vbInformation
End Sub
Private Function GenerateNewSect_ID() As String
    ' Calculate the New Sect_ID value based on existing records
    Dim lastSect_ID As Variant
    Dim newSect_ID As Integer
    
    lastSect_ID = DMax("Sect_ID", "SectorsT", "IsNumeric(Sect_ID) = True")

    If IsNull(lastSect_ID) Then
        ' If the table is empty, set the initial Sect_ID to 1
        newSect_ID = 1
    Else
        ' Increment the last Sect_ID value by 1
        newSect_ID = CInt(lastSect_ID) + 1
    End If

    ' Return the New Sect_ID value
    GenerateNewSect_ID = Format(newSect_ID, "00000")
End Function

Private Sub ResetFormFields()
    ' Clear the form
    Me.SectorNo.Value = ""
    Me.SectorLocation.Value = ""
    Me.Location.Value = ""
    Me.ETD.Value = ""
    Me.ETA.Value = ""
    Me.ATD.Value = ""
    Me.ATA.Value = ""

    ' ... Clear values of other form controls
End Sub
Private Function GetLastRefNo() As Variant
    ' Get the latest AddedTime value from FlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim lastRefNo As Variant

    Set db = CurrentDb
    strSQL = "Select MAX(AddedTime) As LatestAddedTime FROM FlightRequestsT"
    Set rs = db.OpenRecordset(strSQL)

    ' Check If there is a latest AddedTime value
    If Not rs.EOF Then
        ' Get the record With the latest AddedTime value from FlightRequestsT table
        strSQL = "Select RefNo FROM FlightRequestsT WHERE AddedTime = (Select MAX(AddedTime) FROM FlightRequestsT)"
        Set rs = db.OpenRecordset(strSQL)

        ' Check If a record was found
        If Not rs.EOF Then
            lastRefNo = rs.Fields("RefNo").Value
        End If
    End If

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLastRefNo = lastRefNo
End Function

