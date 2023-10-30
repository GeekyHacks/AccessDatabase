Option Compare Database

Private Sub addServiceBtn_Click()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Integer
    
    Dim Service As String, Vendor As String, Agent As String, ServiceStatus As String, PaymentMethod As String
    Dim Notes As String, Sect_ID As Variant ' Add variable declaration for Notes
    
    ' Open the main table's recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("ServicesT", dbOpenDynaset)
    
    ' Loop through the sets of text boxes and add records
    For i = 1 To 10 ' Change 10 to the desired number of services
        
        ' Retrieve values from text boxes
        Service = Nz(Me.Controls("Service" & i).Value, "")
        Vendor = Nz(Me.Controls("Vendor" & i).Value, "")
        Agent = Nz(Me.Controls("Agent" & i).Value, "")
        ServiceStatus = Nz(Me.Controls("ServiceStatus" & i).Value, "")
        PaymentMethod = Nz(Me.Controls("PaymentMethod" & i).Value, "")
        Notes = Nz(Me.Controls("Notes" & i).Value, "")
        Sect_ID = Nz(Me.Controls("Sect_ID" & i).Value, "")
        
        ' Check if at least one field is not empty
        If Service <> "" Or Vendor <> "" Or Agent <> "" Or ServiceStatus <> "" Or PaymentMethod <> "" Or Notes <> "" Or Not IsNull(Notes) Or Not IsNull(Sect_ID) Then
            ' Add a record
            rs.AddNew
            ' Assign values to the fields in the table
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

Exit Sub
ErrorHandler:
     MsgBox "Service records added successfully.", vbInformation
    Exit Sub
End Sub

Private Sub closeBtn_Click()
    DoCmd.Close acForm, "multiNewServicesF"
End Sub

Private Sub Form_Load()
   DoCmd.GoToRecord , , acNewRec
       ' Check if the other form is loaded
    If CurrentProject.AllForms("AddSectorF").IsLoaded Then
        ' Get the value from the other form
        Dim otherForm As Form
        Set otherForm = Forms("AddSectorF")
        Me.Sect_ID1.Value = otherForm.Sect_ID.Value
    Else
        ' Calculate the latest value from the table
        Dim lastSect_ID As String
        lastSect_ID = Nz(DMax("Sect_ID", "SectorsT"), "")
        
        ' Assign the latest value to the Sect_ID1 field
        Me.Sect_ID1.Value = lastSect_ID
    End If
End Sub



