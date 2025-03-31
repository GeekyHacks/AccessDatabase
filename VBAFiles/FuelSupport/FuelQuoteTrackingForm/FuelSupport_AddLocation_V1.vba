Option Compare Database
Option Explicit

Private Sub addLocID()
    ' Generates a unique sequential location ID
    ' Retrieves current maximum ID from Loc_IDT table
    ' Formats New ID As 4-digit number (e.g., "0001")
    ' Updates counter in Loc_IDT table
    ' Sets LocationQuote_ID field With New ID
    Dim RefValue As Integer
    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    RefValue = Nz(DLookup("LocNum", "Loc_IDT"), 0)

    Me.LocationQuote_ID.value = Format(RefValue + 1, "0000")

    Set db = CurrentDb
    Set rst = db.OpenRecordset("Loc_IDT", dbOpenDynaset)

    If Not rst.EOF Then
        'rst.MoveFirst
        rst.Edit
        rst!LocNum = rst!LocNum + 1
        rst.Update
    End If

    rst.Close
    Set rst = Nothing
    Set db = Nothing

End Sub
Private Sub AddLocationBtn_Click()
    ' Main handler For adding New locations To quote requests
    ' Performs validation checks For required fields
    ' Checks For duplicate locations in same request
    ' Creates New record in FuelSupport_RequestedLocationsT_V1
    ' Includes:
    '   - Location ID generation
    '   - User tracking
    '   - Location details
    '   - Status tracking
    '   - Request reference
    ' Automatically increments location number
    ' Clears location field For Next entry
    ' Refreshes locations list display
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim CurrentForm As Form
    Dim newSectorNo As Variant

    addLocID

    Set CurrentForm = Forms("FuelSupport_QuoteRequestF_V1").Form
    If IsNull(Me.LocationQuote_ID) Or IsNull(Me.User) Or IsNull(Me.Location) Or IsNull(Me.LocNo) Or IsNull(Me.LocationQuoteStatus) Or IsNull(Me.RequestRefNo) Then

        MsgBox "Please add missing fields.", vbExclamation
     Exit Sub
    End If


    ' Check If a record With the same SectorNo, Sect_ID, ORIGIN, And Destination already exists
    Set db = CurrentDb
    Set rs = CurrentDb.OpenRecordset("Select Location FROM FuelSupport_RequestedLocationsT_V1 WHERE QuoteRequestRefNo = '" & Me.RequestRefNo.value & "' And Location = '" & Me.Location.value & "'")

    If Not rs.EOF Then
        MsgBox "A record With the same location already exists.", vbExclamation, "Duplicate Record"
        rs.Close
        Set rs = Nothing
        Set db = Nothing
     Exit Sub
    End If

    Set rs = db.OpenRecordset("FuelSupport_RequestedLocationsT_V1", dbOpenDynaset)
    rs.AddNew
    ' Assign values To the fields in the table

    rs.Fields("LocationQuote_ID").value = Me.LocationQuote_ID.value
    rs.Fields("GeneratedBy").value = Me.User.value
    rs.Fields("Location").value = Me.Location.value
    rs.Fields("LocNo").value = Me.LocNo.value
    rs.Fields("LocationQuoteStatus").value = Me.LocationQuoteStatus.value
    rs.Fields("QuoteRequestRefNo").value = Me.RequestRefNo.value



    ' Save the New record
    rs.Update


    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    newSectorNo = Me.LocNo.value + 1
    Me.LocNo.value = newSectorNo
    Me.Location.value = Null

    CurrentForm.Form.Controls("FuelSupport_LocationsListF").Requery

    'Me.Refresh

    ' Display a message indicating successful addition
    MsgBox "Data added successfully."

End Sub

Private Sub Form_Load()
    ' Form initialization handler
    ' Updates user information when form loads
    UpdateUser
End Sub
Private Sub UpdateUser()
    ' Retrieves And displays current user information
    ' Gets user ID from loggedInUser form
    ' Looks up username from UsersT table
    ' Sets User field With current username
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub
