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
    LogToFile "Client email: " & recipientEmail

    ' Construct and sanitize file name
    fileName = SanitizeFileName(Me.Quote_RefNotxt & " " & AIRPORT & " " & clientID & ".pdf")
    LogToFile "File name: " & fileName

    ' Open file dialog to select path
    Dim fd As fileDialog
    Set fd = Application.fileDialog(msoFileDialogFolderPicker)
    Dim filePath As String

    With fd
        .Title = "Select Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1) & "\" & fileName
            LogToFile "Selected file path: " & filePath
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    ' Export report to PDF
    DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
    DoCmd.OutputTo acOutputReport, "FuelSupport_FuelQuote", acFormatPDF, filePath, False
    DoCmd.Close acReport, "FuelSupport_FuelQuote"
    LogToFile "Report exported to: " & filePath

    ' Send email if recipient is found
    If recipientEmail <> "" Then
        Dim olApp As Object
        Dim olNamespace As Object
        Dim olInboxFolder As Object
        Dim olSentItemsFolder As Object
        Dim olOccInboxFolder As Object
        Dim olOccSentItemsFolder As Object
        Dim olMailItem As Object
        Dim olReply As Object
        Dim foundEmail As Boolean
        Dim searchSubject As String
        Dim lastMatchingEmail As Object ' To store the last matching email
        Dim currentUserEmail As String ' To store the current user's email

        searchSubject = CStr(Me.R_RefNo)
        LogToFile "Search subject: " & searchSubject

        Set olApp = CreateObject("Outlook.Application")
        Set olNamespace = olApp.GetNamespace("MAPI")

        ' Get Abdullah's Inbox and Sent Items folders
        Set olInboxFolder = olNamespace.GetDefaultFolder(6) ' Inbox folder
        Set olSentItemsFolder = olNamespace.GetDefaultFolder(5) ' Sent Items folder

        ' Get OCC's Inbox and Sent Items folders
        On Error Resume Next ' Handle potential errors if OCC's mailbox is not accessible
        Set olOccInboxFolder = olNamespace.Folders("occ@tahseenaviation.com").Folders("Inbox")
        Set olOccSentItemsFolder = olNamespace.Folders("occ@tahseenaviation.com").Folders("Sent Items")
        On Error GoTo ErrorHandler ' Reset error handling

        ' Get the current user's email address
        currentUserEmail = olNamespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        LogToFile "Current user email: " & currentUserEmail

        foundEmail = False
        Set lastMatchingEmail = Nothing ' Initialize to Nothing

        ' Search Abdullah's Inbox folder
        LogToFile "Searching Abdullah's Inbox folder..."
        For Each olMailItem In olInboxFolder.Items
            If olMailItem.Class = 43 Then ' olMail
                ' Check if the subject matches
                If InStr(olMailItem.Subject, searchSubject) > 0 Then
                    ' Check if the receiver's email matches the client's primary email
                    If EmailMatchesRecipient(olMailItem.To, recipientEmail, olMailItem) Then
                        ' Store the last matching email
                        Set lastMatchingEmail = olMailItem
                        foundEmail = True
                        LogToFile "Matching email found in Abdullah's Inbox - Subject: " & olMailItem.Subject & ", To: " & olMailItem.To
                    Else
                        LogToFile "Email does not match recipient - To Field: " & olMailItem.To & ", Recipient Email: " & recipientEmail
                    End If
                End If
            End If
        Next olMailItem

        ' Search Abdullah's Sent Items folder if no match found in Inbox
        If Not foundEmail Then
            LogToFile "Searching Abdullah's Sent Items folder..."
            For Each olMailItem In olSentItemsFolder.Items
                If olMailItem.Class = 43 Then ' olMail
                    ' Check if the subject matches
                    If InStr(olMailItem.Subject, searchSubject) > 0 Then
                        ' Check if the receiver's email matches the client's primary email
                        If EmailMatchesRecipient(olMailItem.To, recipientEmail, olMailItem) Then
                            ' Store the last matching email
                            Set lastMatchingEmail = olMailItem
                            foundEmail = True
                            LogToFile "Matching email found in Abdullah's Sent Items - Subject: " & olMailItem.Subject & ", To: " & olMailItem.To
                        Else
                            LogToFile "Email does not match recipient - To Field: " & olMailItem.To & ", Recipient Email: " & recipientEmail
                        End If
                    End If
                End If
            Next olMailItem
        End If

        ' Search OCC's Inbox folder if no match found in Abdullah's folders
        If Not foundEmail And Not olOccInboxFolder Is Nothing Then
            LogToFile "Searching OCC's Inbox folder..."
            For Each olMailItem In olOccInboxFolder.Items
                If olMailItem.Class = 43 Then ' olMail
                    ' Check if the subject matches
                    If InStr(olMailItem.Subject, searchSubject) > 0 Then
                        ' Check if the receiver's email matches the client's primary email
                        If EmailMatchesRecipient(olMailItem.To, recipientEmail, olMailItem) Then
                            ' Store the last matching email
                            Set lastMatchingEmail = olMailItem
                            foundEmail = True
                            LogToFile "Matching email found in OCC's Inbox - Subject: " & olMailItem.Subject & ", To: " & olMailItem.To
                        Else
                            LogToFile "Email does not match recipient - To Field: " & olMailItem.To & ", Recipient Email: " & recipientEmail
                        End If
                    End If
                End If
            Next olMailItem
        End If

        ' Search OCC's Sent Items folder if no match found in OCC's Inbox
        If Not foundEmail And Not olOccSentItemsFolder Is Nothing Then
            LogToFile "Searching OCC's Sent Items folder..."
            For Each olMailItem In olOccSentItemsFolder.Items
                If olMailItem.Class = 43 Then ' olMail
                    ' Check if the subject matches
                    If InStr(olMailItem.Subject, searchSubject) > 0 Then
                        ' Check if the receiver's email matches the client's primary email
                        If EmailMatchesRecipient(olMailItem.To, recipientEmail, olMailItem) Then
                            ' Store the last matching email
                            Set lastMatchingEmail = olMailItem
                            foundEmail = True
                            LogToFile "Matching email found in OCC's Sent Items - Subject: " & olMailItem.Subject & ", To: " & olMailItem.To
                        Else
                            LogToFile "Email does not match recipient - To Field: " & olMailItem.To & ", Recipient Email: " & recipientEmail
                        End If
                    End If
                End If
            Next olMailItem
        End If

        ' If a matching email was found, reply to the last one
        If foundEmail And Not lastMatchingEmail Is Nothing Then
            Set olReply = lastMatchingEmail.ReplyAll
            With olReply
               ' .SentOnBehalfOfName = "occ@tahseenaviation.com"
                .CC = currentUserEmail ' & ";occ@tahseenaviation.com"
                .Attachments.Add filePath
                .HTMLBody = "Dear On Duty, " & "<br><br>" & _
                            "Thank you for your inquiry and for choosing Tahseen Aviation Services." & "<br><br>" & _
                            "Tahseen Fuels Quote #" & Me.Quote_RefNotxt & " for " & AIRPORT & " has been forwarded to your email address. We kindly request that you send us the confirmed flight details to proceed with the order." & "<br><br>" & _
                            "Please do not hesitate to contact us if you have any further fuel requests or inquiries." & "<br><br>" & _
                             "<br>" & .HTMLBody
                If MsgBox("Are you sure you want to send this email?", vbYesNo + vbQuestion) = vbYes Then
                    .Send
                    MsgBox "Reply-all email sent with the attached quote.", vbInformation, "Email Sent"
                    LogToFile "Email sent successfully."
                    ChangeQuoteStatus
                Else
                    MsgBox "Email not sent.", vbInformation, "Operation Cancelled"
                    LogToFile "Email sending cancelled by user."
                End If
            End With
        Else
            MsgBox "No matching email found in the Inbox or Sent Items. Please reply to the email manually with the correct Quote Reference Number.", vbExclamation, "Email Error"
            LogToFile "No matching email found in Inbox or Sent Items."
        End If
    Else
        MsgBox "Client email not found.", vbExclamation, "Email Error"
        LogToFile "Client email not found."
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    LogToFile "Error in EmailQuoteBtn_Click: " & Err.Description
End Sub

' Function to check if the email's To field matches the recipient email
Function EmailMatchesRecipient(emailToField As String, recipientEmail As String, olMailItem As Object) As Boolean
    Dim emailAddress As String
    Dim emailParts As Variant
    Dim i As Integer
    Dim resolvedEmail As String

    ' Initialize the lookup table if it hasn't been initialized yet
    If displayNameLookup Is Nothing Then
        InitializeLookupTable
    End If

    ' Split the To field into individual email addresses
    emailParts = Split(emailToField, ";")

    ' Loop through each email address in the To field
    For i = LBound(emailParts) To UBound(emailParts)
        emailAddress = Trim(emailParts(i)) ' Remove leading/trailing spaces

        ' Log the email address being processed
        LogToFile "Processing email address: " & emailAddress

        ' If the email address is a display name, resolve it to the actual email address
        If InStr(emailAddress, "@") = 0 Then
            ' Resolve the display name using the lookup table
            If displayNameLookup.Exists(emailAddress) Then
                resolvedEmail = displayNameLookup(emailAddress)
                LogToFile "Resolved email address using lookup table: " & resolvedEmail
            Else
                resolvedEmail = emailAddress ' Fallback to the original value
                LogToFile "Display name not found in lookup table: " & emailAddress
            End If
        Else
            resolvedEmail = emailAddress ' Already an email address
            LogToFile "Resolved email address: " & resolvedEmail
        End If

        ' Compare the normalized email address with the recipient email
        If LCase(resolvedEmail) = LCase(recipientEmail) Then
            ' Log the matched email
            LogToFile "Matched email - To Field: " & emailToField & ", Resolved Email: " & resolvedEmail & ", Recipient Email: " & recipientEmail
            EmailMatchesRecipient = True
            Exit Function
        End If
    Next i

    ' If no match is found, return False
    EmailMatchesRecipient = False
End Function

' Function to initialize the lookup table
Private Sub InitializeLookupTable()
    Set displayNameLookup = CreateObject("Scripting.Dictionary")
    
    ' Add display name to email address mappings
    displayNameLookup("Geeky Hacks") = "geekyhacks22@gmail.com"
    displayNameLookup("Deviprasad Shetty") = "devi@example.com"
    displayNameLookup("UCIG Fuel") = "ucig@example.com"
    ' Add more mappings as needed
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

    filePath = "E:\ent\DebugLog.txt" ' Change to a secure location
    fileNumber = FreeFile

    Open filePath For Append As #fileNumber
    Print #fileNumber, Now() & " - " & message
    Close #fileNumber
End Sub