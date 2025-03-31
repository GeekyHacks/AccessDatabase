Option Compare Database
Option Explicit

Private Sub AddQuotebtn_Click()
    ' Primary handler For initiating quote creation
    ' Sets focus To the fuel quote form
    ' Calls RenderLocationRecord To populate form With location data
    ' Includes basic error handling
    On Error Goto ErrHandler

        Dim selectedSect_ID As Long
        Forms("FuelSupport_QuoteRequestF_V1").Controls("FuelSupport_AddFuelQuote").SetFocus
        RenderLocationRecord
        Me.Refresh
     Exit Sub

 ErrHandler:
        MsgBox "Error in AddQuotebtn_Click: " & Err.Description, vbCritical
End Sub

Private Sub RenderLocationRecord()
    ' Core Function that populates quote form With location-specific data
    ' Retrieves:
    '   - Location details from RequestedLocationsT_V1 table
    '   - User information from UsersT table
    '   - Client data from parent form
    ' Validates required fields before processing
    ' Sets values in target form controls:
    '   - User name
    '   - Request reference
    '   - Airport location
    '   - Client information
    ' Includes comprehensive error handling
    On Error Goto ErrHandler

        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim selectedLoc As Long
        Dim QuoteForm As Form
        Dim Client As String
        Dim UserID As Integer
        Dim CLientID As Integer

        Client = Forms("FuelSupport_QuoteRequestF_V1").Controls("Customer").value
        CLientID = Forms("FuelSupport_QuoteRequestF_V1").Controls("Customer_IDtxt").value
        If IsNull(Me.LocationQuote_IDtxt.value) Then
            MsgBox "Please Select a valid Sector ID.", vbExclamation
         Exit Sub
        End If

        selectedLoc = Me.LocationQuote_IDtxt.value

        ' Validate UserID
        If Forms!loggedInUser Is Nothing Or IsNull(Forms!loggedInUser!txtUserID) Then
            MsgBox "User ID Not found.", vbExclamation
         Exit Sub
        End If

        UserID = Forms!loggedInUser!txtUserID

        ' Reference the subform correctly
        Set QuoteForm = Forms("FuelSupport_QuoteRequestF_V1").Controls("FuelSupport_AddFuelQuote").Form

        Set db = CurrentDb
        Set rs = db.OpenRecordset("Select Location, QuoteRequestRefNo FROM FuelSupport_RequestedLocationsT_V1 WHERE LocationQuote_ID = " & selectedLoc, dbOpenSnapshot)

        If rs.EOF Then
            MsgBox "No record found For location ID: " & selectedLoc, vbInformation
            Goto Cleanup
            End If

            With QuoteForm
                .Controls("Usertxt").value = DLookup("UserName", "UsersT", "ID =" & UserID)
                .Controls("R_RefNo").value = Nz(rs("QuoteRequestRefNo").value, "")
                .Controls("AIRPORT").value = Nz(rs("Location").value, "")
                .Controls("Client").value = Client
                .Controls("ClientIDtxt").value = CLientID
                .Controls("AIRPORT").SetFocus
            End With

 Cleanup:
            If Not rs Is Nothing Then rs.Close
                Set rs = Nothing
                Set db = Nothing
                Set QuoteForm = Nothing
             Exit Sub

 ErrHandler:
                MsgBox "Error in RenderLocationRecord: " & Err.Description, vbCritical
                Resume Cleanup
End Sub
Private Sub QuoteRefNotxt_DblClick(Cancel As Integer)
    PrintQuote_Click
End Sub

Private Sub PrintQuote_Click()
    ' Handles quote report generation And display
    ' Creates filtered report based on quote reference
    ' Falls back To report generator form If no reference exists
    ' Opens report in preview mode
    Dim strWhere As String
    Dim qouteRefno As Variant

    ' Get the value from the QuoteRefnotxt control
    qouteRefno = Me.QuoteRefnotxt.value

    ' Check If the QuoteRefnotxt is Not null
    If Not IsNull(qouteRefno) Then
        ' Build the filter string For the report
        strWhere = "Quote_RefNo = '" & qouteRefno & "'"

        ' Open the report With the filter
        DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
    Else
        ' If QuoteRefnotxt is null, open the report generator form
        DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
    End If
End Sub
