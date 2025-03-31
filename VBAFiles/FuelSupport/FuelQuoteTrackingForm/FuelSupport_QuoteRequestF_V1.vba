Option Compare Database
Option Explicit
Private Sub cmdGenerate_Click()
    ' Generates a unique random quote request reference number
    ' Format: 0000YY (4-digit random number + 2-digit year)
    ' Ensures uniqueness by checking against existing requests
    ' Uses a Loop To regenerate If duplicate is found
    ' Sets the generated number in the RefNo form field
    Dim rst As DAO.Recordset
    Dim randomRefValue As Integer
    Dim SelectedRefNo As String
    Dim isUnique As Boolean

    isUnique = False ' Initialize isUnique To False

    Do
        ' Generate a random number between 1 And 9999
        randomRefValue = Int((9999 - 1 + 1) * Rnd + 1)

        ' Generate the reference number
        SelectedRefNo = Format(randomRefValue, "0000") & Format(date, "yy")

        ' Check If the generated reference number already exists in the FuelSupport_QuoteRequestsT_V1 table
        Set rst = CurrentDb.OpenRecordset("Select QuoteRequestRefNo FROM FuelSupport_QuoteRequestsT_V1 WHERE QuoteRequestRefNo = '" & SelectedRefNo & "'", dbOpenSnapshot)
        If rst.EOF Then
            ' If the generated reference number does Not exist, assign it To the form field
            Me.RefNo.value = SelectedRefNo

            ' Set the flag To indicate that a unique reference number has been generated
            isUnique = True
        End If

        rst.Close
        Set rst = Nothing

    Loop Until isUnique

End Sub

Private Sub newRequestBtn_Click()
    ' Prepares form For New request entry
    ' Generates reference number If empty
    ' Sets focus To Customer field For data entry

    If IsNull(Me.RefNo) Or Me.RefNo = "" Then
        cmdGenerate_Click
    End If
    ' Set focus To the first field For data entry
    Me.Customer.SetFocus
End Sub

Private Sub AddQuoteRequestBtn_Click()
    ' Validates And saves New quote request To database
    ' Checks all required fields are completed
    ' Creates New record in FuelSupport_QuoteRequestsT_V1
    ' Includes timestamp For last modification
    ' Displays success message And resets form
    ' Loads the newly created record For verification
    Dim SelectedRefNo As Variant
    Dim CurrentForm As Form
    newRequestBtn_Click
    ' Check If all the required fields are filled
    If IsNull(Me.RefNo.value) Or IsNull(Me.ReqDate.value) Or IsNull(Me.RequiredDate.value) Or IsNull(Me.Operator.value) Or IsNull(Me.FlightType.value) Or IsNull(Me.RequestedLocationsList.value) Or IsNull(Me.Customer.value) Or IsNull(Me.RequestStatus.value) Or IsNull(Me.Customer_IDtxt.value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
     Exit Sub
    Else
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim currentTimestamp As Variant

        Set db = CurrentDb
        Set rs = db.OpenRecordset("FuelSupport_QuoteRequestsT_V1", dbOpenDynaset)
        rs.AddNew

        ' Assign values To the fields in the table
        With rs
            .Fields("QuoteRequestRefNo").value = Me.RefNo.value
            ' .Fields("ReqNo").value = Me.ReqNo.value    'Not needed
            .Fields("ReqDate").value = Me.ReqDate.value
            .Fields("RequiredDate").value = Me.RequiredDate.value
            .Fields("FlightType").value = Me.FlightType.value
            .Fields("RequestStatus").value = Me.RequestStatus.value
            .Fields("Notes").value = Me.Notes.value
            .Fields("RequestedLocationsList").value = Me.RequestedLocationsList.value
            .Fields("CustomerName").value = Me.Customer.value
            .Fields("Customer_ID").value = Me.Customer_IDtxt.value
            .Fields("Operator").value = Me.Operator.value
            .Fields("Urgent").value = Me.Urgent.value
            ' Store the current timestamp For the record
            currentTimestamp = Now
            .Fields("LastModifiedTimestamp").value = currentTimestamp
            .Update
        End With

        ' Display a message indicating successful addition
        MsgBox "Quote Request added successfully, Now add locations As per the request"

        rs.Close
        Set rs = Nothing
        Set db = Nothing

        Me.Undo

        ' Display the last added record
        Dim MaxID As Long
        MaxID = DMax("ID", "FuelSupport_QuoteRequestsT_V1")
        Me.RecordSource = "Select CustomerName, Customer_ID, QuoteRequestRefNo, ReqDate, RequiredDate, FlightType, RequestedLocationsList, RequestStatus, Notes,GeneratedBy, Operator FROM FuelSupport_QuoteRequestsT_V1 WHERE ID = " & MaxID & ";"
        Me.Requery

        ' Get the selected RefNo value
        SelectedRefNo = Me.RefNo.value

        ' Reference the subform within the main form
        Set CurrentForm = Forms("FuelSupport_QuoteRequestF_V1").Form

        ' Set refno To the subform
        ' CurrentForm.Controls("AddRequestedLocationsF").Controls("RequestRefNo").Value = SelectedRefNo

     Exit Sub
    End If
End Sub

Private Sub CloseBtn_Click()
    DoCmd.Close acForm, "FuelSupport_QuoteRequestF_V1"
End Sub
Private Sub Customer_AfterUpdate()
    ' Synchronizes customer ID when customer name is selected
    ' Looks up ID from CustomersT table based on selected name
    ' Updates Customer_IDtxt field With matching ID
    Dim selectedValue As Variant
    selectedValue = Me.Customer.value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("CustomersT", dbOpenSnapshot)

        rs.FindFirst "Customer = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.Customer_IDtxt.value = rs("ID").value
        Else
            Me.Customer_IDtxt.value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.Customer_IDtxt.value = Null
    End If
End Sub
Private Sub Customer_Enter()
    Customer_AfterUpdate
End Sub

Private Sub FullBriefBtn_Click()
    Dim strWhere As String
    Dim InLeng As Long
    If Not IsNull(Me.RefNo) Then
        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNo & "*"") And "
    End If

    InLeng = Len(strWhere) - 5
    If InLeng <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, InLeng)
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_FullExternalReport_V12", acViewPreview, , strWhere
    End If
End Sub

Private Sub RefreshBtn_Click()
    Me.Requery
    Me.Refresh
End Sub

Private Sub ExtBriefBtn_Click()
    Dim strWhere As String
    Dim InLeng As Long
    If Not IsNull(Me.RefNo) Then
        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNo & "*"") And "
    End If

    InLeng = Len(strWhere) - 5
    If InLeng <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, InLeng)
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_ExternalReport_New", acViewPreview, , strWhere
    End If
End Sub

Private Sub FlightBriefBtn_Click()
    Dim strWhere As String
    Dim InLeng As Long

    If Not IsNull(Me.RefNo) Then
        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNo & "*"") And "
    End If

    InLeng = Len(strWhere) - 5
    If InLeng <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, InLeng)
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_FullInternalReport_V12", acViewPreview, , strWhere
    End If
End Sub

Private Sub Form_Load()
    Dim fw As New clFormWindow

    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
    LoadLastRequestBtn_Click
End Sub

Private Sub LoadLastRequestBtn_Click()
    Dim MaxID As Long
    Dim strSQL As String

    ' Get the maximum ID from the table
    MaxID = Nz(DMax("ID", "FuelSupport_QuoteRequestsT_V1"), 0)

    ' Construct the RecordSource query
    strSQL = "Select CustomerName, Customer_ID, QuoteRequestRefNo, ReqDate, Urgent, RequiredDate, FlightType, RequestedLocationsList, RequestStatus, Notes, GeneratedBy, Operator FROM FuelSupport_QuoteRequestsT_V1 WHERE ID = " & MaxID & ";"

    ' Set the RecordSource Property
    Me.RecordSource = strSQL

    ' Requery the form
    Me.Requery
    Me.RefNo2.value = Me.RefNo.value
    Me.Controls("AddRequestedLocationsF").Controls("RequestRefNo").value = Me.RefNo.value
    FilterLocationsList
End Sub

Private Sub RequiredDate_AfterUpdate()
    If Me.RequiredDate.value = Me.ReqDate.value Then
        Me.Urgent.value = True
    End If
End Sub

Private Sub RequiredDate_Click()
    InputDateField RequiredDate, "Select a date To use this on your form"
End Sub

Public Sub TriggerAfterUpdate()
    Call RefNo2_AfterUpdate
End Sub

Private Sub RefNo2_AfterUpdate()
    ' Filters form And subform when reference number changes
    ' Updates record source To show selected request
    ' Synchronizes With locations subform
    Dim SelectedRefNo As Variant
    Dim CurrentForm As Form
    FilterLocationsList
    ' Get the selected RefNo value
    SelectedRefNo = Me.RefNo2.value

    ' Reference the subform within the main form
    '  Set CurrentForm = Forms("FuelSupport_QuoteRequestF_V1").Form

    ' Set the record source based on the selected RefNo
    Me.RecordSource = "Select CustomerName, Customer_ID, QuoteRequestRefNo, ReqDate, RequiredDate, FlightType, Urgent, RequestedLocationsList, RequestStatus, Notes,GeneratedBy, Operator FROM FuelSupport_QuoteRequestsT_V1 WHERE QuoteRequestRefNo = '" & SelectedRefNo & "';"
    ' Set refno To the subform0
    Me.Controls("AddRequestedLocationsF").Controls("RequestRefNo").value = SelectedRefNo
    Me.Requery
    On Error Goto 0
        ' Reset dropdowns, reselect, And apply filters
        ' Reset_DropDowns
        ' RefNo_reselect
        ' activeFilter

        ' Requery And refresh the form
        Me.Requery

End Sub

Private Sub FilterLocationsList()
    ' Applies filter To locations subform
    ' Shows only locations matching current request reference
    Dim SelectedRefNo As Variant
    Dim FuelForm As Form

    Set FuelForm = Me.Controls("FuelSupport_LocationsListF").Form

    ' Wait For the form To be loaded
    SelectedRefNo = Me.RefNo2.value

    Me.FuelSupport_LocationsListF.Form.Filter = "QuoteRequestRefNo = '" & SelectedRefNo & "'"
    Me.FuelSupport_LocationsListF.Form.FilterOn = True

    Me.Requery
End Sub

Public Sub ReqDate_Click()
    InputDateField ReqDate, "Select a date To use this on your form"
End Sub

Private Sub ResetBtn_Click()
    On Error Goto ErrorHandler

        Me.Undo
        DoCmd.CancelEvent
        DoCmd.GoToRecord , , acNewRec

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & Err.Description, vbExclamation
End Sub
