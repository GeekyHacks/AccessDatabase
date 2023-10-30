
Option Compare Database

Private Sub AddServicesBtn_Click()
    DoCmd.OpenForm "multiNewServicesF", WindowMode:=acDialog
End Sub

Private Sub closeBtn_Click()
    ' Check if there are any unsaved changes
    If Me.Dirty Then
        ' Prompt the user to save changes
        Dim response As Integer
        response = MsgBox("Do you want to save the changes?", vbQuestion + vbYesNoCancel)
        
        Select Case response
            Case vbYes
                ' Save the changes
                Me.Dirty = False ' Save the changes made on the form
                DoCmd.Close acForm, Me.Name ' Close the form
            Case vbNo
                ' Discard the changes
                Me.Undo ' Undo the changes made on the form
                DoCmd.Close acForm, Me.Name ' Close the form
            Case vbCancel
                ' Cancel the close operation
                Exit Sub
        End Select
    Else
        ' No unsaved changes, close the form
        DoCmd.Close acForm, Me.Name
    End If
End Sub
Private Sub newSectorBtn_Click()
    ' Check If there are any unsaved changes
    If Me.Dirty Then
        Me.Undo ' Undo any changes To cancel the current record
    End If

    ' Clear the form fields To prepare For a New record
    ResetFormFields

    ' Set focus To the first field For data entry
    Me.SectorNo.SetFocus
End Sub


Private Function ValidateFormFields() As Boolean
    ' Check if any required fields are empty
    If Me.SectorNo.Value = "" Then
        MsgBox "Sector Number is required.", vbExclamation
        Me.SectorNo.SetFocus
        ValidateFormFields = False
        Exit Function
    End If
    
    If Me.SectorLocation.Value = "" Then
        MsgBox "Sector Location is required.", vbExclamation
        Me.SectorLocation.SetFocus
        ValidateFormFields = False
        Exit Function
    End If
    
    ' Add more field validations as needed
    
    ' If all validations pass, return True
    ValidateFormFields = True
End Function

Private Sub addSectorBtn_Click()
    ' Check if all required fields are filled in
    If Not ValidateFormFields() Then
        Exit Sub
    End If

    ' Open the main table's recordset
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
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
        ' Generate a New Sect_ID And assign it To the form control
    Me.Sect_ID.Value = GenerateNewSect_ID()
    Me.RefNo.Value = GetLastRefNo()
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

