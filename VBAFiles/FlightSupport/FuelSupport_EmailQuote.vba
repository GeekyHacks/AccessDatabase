Private Sub EmailQuoteBtn_Click()
    On Error GoTo ErrorHandler

    ' Validate input
    If IsNull(Me.Quote_RefNotxt) Or Me.Quote_RefNotxt = "" Then
        MsgBox "Quote Reference Number is required.", vbExclamation
        Exit Sub
    End If

    ' Sanitize input
    Dim sanitizedQuoteRefNo As String
    sanitizedQuoteRefNo = Replace(Me.Quote_RefNotxt, "'", "''")

    ' Construct filter
    Dim strWhere As String
    strWhere = "([Quote_RefNo] Like '*" & sanitizedQuoteRefNo & "*') And "
    strWhere = Left$(strWhere, Len(strWhere) - 5)

    ' Apply filter to the form
    Me.Filter = strWhere
    Me.FilterOn = True

    ' Extract values from the report
    Dim AIRPORT As String
    Dim clientID As String
    Dim recipientEmail As String
    Dim fileName As String
    Dim rs As DAO.Recordset

    ' Use parameterized query to prevent SQL injection
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.CreateQueryDef("", "SELECT DISTINCT Airport FROM FuelSupport_FuelQuotationsT WHERE [Quote_RefNo] Like @QuoteRefNo")
    qdf.Parameters("@QuoteRefNo").Value = "*" & Me.Quote_RefNotxt & "*"
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    AIRPORT = ""
    Do While Not rs.EOF
        AIRPORT = AIRPORT & Left(rs!AIRPORT, 4) & "_"
        rs.MoveNext
    Loop

    ' Remove trailing underscores
    If Len(AIRPORT) > 0 Then AIRPORT = Left(AIRPORT, Len(AIRPORT) - 1)

    rs.Close
    Set rs = Nothing

    ' Get client email
    clientID = Me.ClientIDtxt.Value
    recipientEmail = Nz(DLookup("PrimaryEmail", "CustomersT", "ID = " & clientID), "")

    ' Construct and sanitize file name
    fileName = SanitizeFileName(Me.Quote_RefNotxt & " " & AIRPORT & " " & clientID & ".pdf")

    ' Open file dialog to select path
    Dim fd As fileDialog
    Set fd = Application.fileDialog(msoFileDialogFolderPicker)
    Dim filePath As String

    With fd
        .Title = "Select Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1) & "\" & fileName
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    ' Export report to PDF
    DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
    DoCmd.OutputTo acOutputReport, "FuelSupport_FuelQuote", acFormatPDF, filePath, False
    DoCmd.Close acReport, "FuelSupport_FuelQuote"

    ' Send email if recipient is found
    If recipientEmail <> "" Then
        Dim olApp As Object
        Dim olNamespace As Object
        Dim olFolder As Object
        Dim olMailItem As Object
        Dim olReply As Object
        Dim foundEmail As Boolean
        Dim searchSubject As String
        Dim lastMatchingEmail As Object ' To store the last matching email

        searchSubject = CStr(Me.R_RefNo)

        Set olApp = CreateObject("Outlook.Application")
        Set olNamespace = olApp.GetNamespace("MAPI")
        Set olFolder = olNamespace.GetDefaultFolder(6) ' Inbox folder

        foundEmail = False
        Set lastMatchingEmail = Nothing ' Initialize to Nothing

        ' Loop through emails in the Inbox to find the last matching email
        For Each olMailItem In olFolder.Items
            If olMailItem.Class = 43 Then ' olMail
                If InStr(olMailItem.Subject, searchSubject) > 0 Then
                    ' Store the last matching email
                    Set lastMatchingEmail = olMailItem
                    foundEmail = True
                End If
            End If
        Next olMailItem

        ' If a matching email was found, reply to the last one
        If foundEmail And Not lastMatchingEmail Is Nothing Then
            Set olReply = lastMatchingEmail.ReplyAll
            With olReply
            ' .SentOnBehalfOfName = "occ@tahseenaviation.com"
                .CC = "Abdullah@tahseenaviation.com" ' Replace with dynamic value
                .Attachments.Add filePath
                .HTMLBody = "Dear On Duty, " & "<br><br>" & _
                            "Thank you for your inquiry and for choosing Tahseen Aviation Services." & "<br><br>" & _
                            "Tahseen Fuels Quote #" & Me.Quote_RefNotxt & " for " & AIRPORT & " has been forwarded to your email address. We kindly request that you send us the confirmed flight details to proceed with the order." & "<br><br>" & _
                            "Please do not hesitate to contact us if you have any further fuel requests or inquiries." & "<br><br>" & _
                            "Kind Regards," & "<br>" & .HTMLBody
                If MsgBox("Are you sure you want to send this email?", vbYesNo + vbQuestion) = vbYes Then
                    .Send
                End If
            End With

            MsgBox "Reply-all email sent with the attached quote.", vbInformation, "Email Sent"
        Else
            MsgBox "No matching email found in the Inbox. Please reply to the email manually with the correct Quote Reference Number.", vbExclamation, "Email Error"
        End If
    Else
        MsgBox "Client email not found.", vbExclamation, "Email Error"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    LogToFile "Error in EmailQuoteBtn_Click: " & Err.Description
End Sub

' Function to sanitize file names
Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer

    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i

    SanitizeFileName = fileName
End Function

' Function to log messages securely
Private Sub LogToFile(message As String)
    Dim filePath As String
    Dim fileNumber As Integer

    filePath = "C:\SecureLogs\DebugLog.txt" ' Change to a secure location
    fileNumber = FreeFile

    Open filePath For Append As #fileNumber
    Print #fileNumber, Now() & " - " & message
    Close #fileNumber
End Sub

