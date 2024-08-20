Option Compare Database

Private Sub closeBtn_Click()
    DoCmd.Close acForm, "ReportGeneratorF"
End Sub

Private Sub GenerateINReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([Customer] Like ""*" & Me.CustomerName & "*"") And "
    End If

    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([Customer] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If

    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([ReqDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([ReqDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "InternalReport", acViewPreview, , strWhere
    End If


End Sub

Private Sub GenerateOUTReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"

    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([Customer] Like ""*" & Me.CustomerName & "*"") And "
    End If
    If Not IsNull(Me.RefNoList) Then

        strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
    End If

    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([ReqDate] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If
    If Not IsNull(Me.EndDate) Then
        strWhere = strWhere & "([ReqDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "ExternalReport", acViewPreview, , strWhere
    End If

End Sub
Dim xhr As Object
Dim url As String
Dim response As String

' Construct the URL For the ASP.NET web page
url = "https://tahseenaviation.sharepoint.com/Lists/OVRF_RequestT/AllItems.aspx"

' Create a New XMLHttpRequest object
Set xhr = CreateObject("MSXML2.XMLHTTP")

' Send a Get request To the web page
xhr.Open "Get", url, False
xhr.send

' Get the response HTML
response = xhr.responseText

' Display the response HTML
MsgBox response

' Release the resources
Set xhr = Nothing


















Dim xhr As Object
Dim url As String
Dim response As String
Dim json As Object
Dim lastRecord As Object

' Construct the URL For the SharePoint list's REST API endpoint
url = "https://tahseenaviation.sharepoint.com/_api/web/lists/getbytitle('OVRF_RequestT')/items?$orderby=ID desc&$top=1"

' Create a New XMLHttpRequest object
Set xhr = CreateObject("MSXML2.XMLHTTP")

' Send a Get request To the SharePoint REST API
xhr.Open "Get", url, False
xhr.setRequestHeader "Accept", "application/json;odata=verbose"
xhr.send

' Get the response JSON
response = xhr.responseText

' Parse the JSON response
Set json = JsonConverter.ParseJson(response)

' Retrieve the last record from the response
Set lastRecord = json("d")("results")(1)

' Access the desired fields from the last record
Dim lastRecordID As String
lastRecordID = lastRecord("ID")
' Access other fields As needed

' Release the resources
Set xhr = Nothing

' Display the last record ID
MsgBox "Last Record ID: " & lastRecordID














Private Sub GenerateINReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"
    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else

        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhere
        Me.FilterOn = True

        If Not IsNull(Me.CustomerName) Then

            strWhere = strWhere & "([Client] Like ""*" & Me.CustomerName & "*"") And "
        End If

        If Not IsNull(Me.OperatorListCb) Then

            strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
        End If
        If Not IsNull(Me.RefNoList) Then

            strWhere = strWhere & "([RefNo] Like ""*" & Me.RefNoList & "*"") And "
        End If

        If Not IsNull(Me.StartDate) Then
            strWhere = strWhere & "([D_Date] >= " & Format(Me.StartDate, conJetDate) & ") And "
        End If

        If Me.lblEndDate.Caption = "Return Date" Then
            If Not IsNull(Me.EndDate) Then
                strWhere = strWhere & "([R_Date] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
                DoCmd.OpenReport "OVRF_REPORT", acViewPreview, , strWhere
            End If
        End If

        If Me.lblEndDate.Caption = "OPS Date" Then
            If Not IsNull(Me.EndDate) Then
                strWhere = strWhere & "([OPSDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
                DoCmd.OpenReport "OPS_OVRF_REPORT", acViewPreview, , strWhere
            End If
        End If
    End If


End Sub






Private Sub GenerateINReportBtn_Click()
    Dim strWehre As String
    Dim InLeng As Long

    Const conJetDate = "\#mm\/dd\/yyyy\#"
    If Not IsNull(Me.PermitNo) Then

        strWhere = strWhere & "([PermitNo] Like ""*" & Me.PermitNo & "*"") And "
    End If
    If Not IsNull(Me.CustomerName) Then

        strWhere = strWhere & "([Client] Like ""*" & Me.CustomerName & "*"") And "
    End If

    If Not IsNull(Me.OperatorListCb) Then

        strWhere = strWhere & "([Operator] Like ""*" & Me.OperatorListCb & "*"") And "
    End If
    If Not IsNull(Me.ReqNoList) Then

        strWhere = strWhere & "([RequestNo] Like ""*" & Me.ReqNoList & "*"") And "
    End If

    If Not IsNull(Me.StartDate) Then
        strWhere = strWhere & "([D_Date] >= " & Format(Me.StartDate, conJetDate) & ") And "
    End If

    If Me.lblEndDate.Caption = "Return Date" Then
        If Not IsNull(Me.EndDate) Then
            strWhere = strWhere & "([R_Date] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
        End If
    End If

    If Me.lblEndDate.Caption = "OPS Date" Then
        If Not IsNull(Me.EndDate) Then
            strWhere = strWhere & "([OPSDate] < " & Format(Me.EndDate + 1, conJetDate) & ") And "
        End If
    End If

    IngLen = Len(strWhere) - 5
    If IngLen <= 0 Then
        MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
    Else
        strWhere = Left$(strWhere, IngLen)
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "OVRF_REPORT", acViewPreview, , strWhere
    End If

End Sub



    
Private Sub RefNo1_GotFocus()
    Me.RefNo1 = Me.CAMA.Column(1)
    
    End Sub
    Private Sub RefNo2_GotFocus()
    Me.RefNo2 = Me.RefNo1
    End Sub
    Private Sub RefNo3_GotFocus()
    Me.RefNo3 = Me.RefNo1
    End Sub
    Private Sub RefNo4_GotFocus()
    Me.RefNo4 = Me.RefNo1
    End Sub
    Private Sub RefNo5_GotFocus()
    Me.RefNo5 = Me.RefNo1
    End Sub


        Private Sub UpdateACinfo(ACRegCb As MSForms.ComboBox, acTypeControl As String, operatorControl As String, customerControl As String)
            Dim selectedValue As Variant
               selectedValue = ACRegCb.Value ' Get the selected value from the combo box
           
               If Not IsNull(selectedValue) Then
                   Dim rs As DAO.Recordset
                   Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)
           
                   rs.FindFirst "ACReg = '" & selectedValue & "'"
                   If Not rs.NoMatch Then
                       Me.Controls(acTypeControl).Value = rs("ACType").Value
              
                       'Me.ACMTOW.Value = rs("ACMTOW").Value
                   Me.Controls(operatorControl).Value = rs("Operator").Value
                       Me.Controls(customerControl).Value = rs("Customer").Value
                   Else
                       Me.Controls(acTypeControl).Value = Null
                       'Me.ACMTOW.Value = Null
                       Me.Controls(operatorControl).Value = Null
                       Me.Controls(customerControl).Value = Null
                   End If
           
                   rs.Close
                   Set rs = Nothing
               Else
                   'Me.acTypeControl.Value = Null
                  'Me.ACMTOW.Value = Null
                   'Me.Operator1.Value = Null
                  ' Me.Customer1.Value = Null
                   Exit Sub
               End If
           
               Exit Sub
           
           End Sub