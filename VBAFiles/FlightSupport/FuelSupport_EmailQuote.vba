Private Sub ChangeQuoteStatus()
    On Error Goto ErrorHandler
        Dim FatherForm As Form
        Dim db As DAO.Database
        Dim rsX As DAO.Recordset
        Dim rsY As DAO.Recordset
        Dim strSQL As String
        Dim QRequestRefNo As String ' Assuming QRequestRefNo is a string; change To Long If it's a number

        ' Initialize database And recordsets
        Set db = CurrentDb
        Set FatherForm = Forms("FuelSupport_QuoteRequestF_V1").Form



        ' Get the QuoteRequestRefNo value (replace this With your actual logic To Get the value)
        QRequestRefNo = Me.R_RefNo ' Example: Get the value from a form control

        ' Open Table Y (FuelSupport_FuelQuotationsT) With a filter For the specific QuoteRequestRefNo
        Set rsY = db.OpenRecordset("Select * FROM FuelSupport_FuelQuotationsT WHERE QuoteRequestRefNo = '" & QRequestRefNo & "'", dbOpenDynaset)

        ' Check If Table Y has records
        If rsY.EOF And rsY.BOF Then
            MsgBox "No records found in Table Y For QuoteRequestRefNo: " & QRequestRefNo, vbExclamation
         Exit Sub
        End If

        ' Loop through Table Y And update Table X For matching locations
        rsY.MoveFirst
        Do While Not rsY.EOF
            ' Open Table X With a filter For the current location And QuoteRequestRefNo
            strSQL = "Select * FROM FuelSupport_RequestedLocationsT_V1 WHERE QuoteRequestRefNo = '" & QRequestRefNo & "' And Location = '" & rsY!AIRPORT & "'"
            Set rsX = db.OpenRecordset(strSQL, dbOpenDynaset)

            ' If a matching record is found, update Table X
            If Not rsX.EOF And Not rsX.BOF Then
                rsX.MoveFirst
                Do While Not rsX.EOF
                    rsX.Edit
                    rsX!LocationQuoteStatus = "Sent"
                    rsX!QuoteRefNo = rsY!Quote_RefNo '' Update values in Table X With QuoteStatus from Table Y
                    rsX.Update
                    rsX.MoveNext
                Loop
            End If

            ' Close the Table X recordset For the current location
            rsX.Close
            Set rsX = Nothing

            ' Move To the Next record in Table Y
            rsY.MoveNext
        Loop

        ' Clean up
        rsY.Close
        Set rsY = Nothing
        Set db = Nothing

        MsgBox "Table X has been updated successfully For QuoteRequestRefNo: " & QRequestRefNo, vbInformation
     Exit Sub
        FatherForm.Controls("FuelSupport_LocationsListF").Form.Requery

 ErrorHandler:
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        ' Clean up in Case of error
        If Not rsX Is Nothing Then rsX.Close
            If Not rsY Is Nothing Then rsY.Close
                Set rsX = Nothing
                Set rsY = Nothing
                Set db = Nothing
End Sub
Private Sub EmailQuoteBtn_Click()
    On Error Goto ErrorHandler

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

        ' Apply filter To the form
        Me.Filter = strWhere
        Me.FilterOn = True

        ' Extract values from the report
        Dim AIRPORT As String
        Dim clientID As String
        Dim recipientEmail As String
        Dim fileName As String
        Dim rs As DAO.Recordset

        ' Use parameterized query To prevent SQL injection
        Dim qdf As DAO.QueryDef
        Set qdf = CurrentDb.CreateQueryDef("", "Select DISTINCT Airport FROM FuelSupport_FuelQuotationsT WHERE [Quote_RefNo] Like @QuoteRefNo")
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

            ' Construct And sanitize file name
            fileName = SanitizeFileName(Me.Quote_RefNotxt & " " & AIRPORT & " " & clientID & ".pdf")

            ' Open file dialog To Select path
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

            ' Export report To PDF
            DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
            DoCmd.OutputTo acOutputReport, "FuelSupport_FuelQuote", acFormatPDF, filePath, False
            DoCmd.Close acReport, "FuelSupport_FuelQuote"

            ' Send email If recipient is found
            If recipientEmail <> "" Then
                Dim olApp As Object
                Dim olNamespace As Object
                Dim olInboxFolder As Object
                Dim olSentItemsFolder As Object
                Dim olMailItem As Object
                Dim olReply As Object
                Dim foundEmail As Boolean
                Dim searchSubject As String
                Dim lastMatchingEmail As Object ' To store the last matching email
                Dim currentUserEmail As String ' To store the current user's email

                searchSubject = CStr(Me.R_RefNo)

                Set olApp = CreateObject("Outlook.Application")
                Set olNamespace = olApp.GetNamespace("MAPI")
                Set olInboxFolder = olNamespace.GetDefaultFolder(6) ' Inbox folder
                Set olSentItemsFolder = olNamespace.GetDefaultFolder(5) ' Sent Items folder

                ' Get the current user's email address
                currentUserEmail = olNamespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                LogToFile "Current user email: " & currentUserEmail

                foundEmail = False
                Set lastMatchingEmail = Nothing ' Initialize To Nothing

                ' Search Inbox folder
                For Each olMailItem In olInboxFolder.Items
                    If olMailItem.Class = 43 Then ' olMail
                        ' Check If the subject matches And the receiver's email matches the client's primary email
                        If InStr(olMailItem.Subject, searchSubject) > 0 And _
                            EmailMatchesRecipient(olMailItem.To, recipientEmail) Then
                            ' Store the last matching email
                            Set lastMatchingEmail = olMailItem
                            foundEmail = True
                        End If
                    End If
                Next olMailItem

                ' Search Sent Items folder If no match found in Inbox
                If Not foundEmail Then
                    For Each olMailItem In olSentItemsFolder.Items
                        If olMailItem.Class = 43 Then ' olMail
                            ' Check If the subject matches And the receiver's email matches the client's primary email
                            If InStr(olMailItem.Subject, searchSubject) > 0 And _
                                EmailMatchesRecipient(olMailItem.To, recipientEmail) Then
                                ' Store the last matching email
                                Set lastMatchingEmail = olMailItem
                                foundEmail = True
                            End If
                        End If
                    Next olMailItem
                End If

                ' If a matching email was found, reply To the last one
                If foundEmail And Not lastMatchingEmail Is Nothing Then
                    Set olReply = lastMatchingEmail.ReplyAll
                    With olReply
                        ' .SentOnBehalfOfName = "occ@tahseenaviation.com"
                        .CC = currentUserEmail' & ";occ@tahseenaviation.com"
                        .Attachments.Add filePath
                        .HTMLBody = "Dear On Duty, " & "<br><br>" & _
                        "Thank you For your inquiry And For choosing Tahseen Aviation Services." & "<br><br>" & _
                        "Tahseen Fuels Quote #" & Me.Quote_RefNotxt & " For " & AIRPORT & " has been forwarded To your email address. We kindly request that you send us the confirmed flight details To proceed With the order." & "<br><br>" & _
                        "Please Do Not hesitate To contact us If you have any further fuel requests Or inquiries." & "<br><br>" & _
                        "<br>" & .HTMLBody
                        If MsgBox("Are you sure you want To send this email?", vbYesNo + vbQuestion) = vbYes Then
                            .Send
                            MsgBox "Reply-all email sent With the attached quote.", vbInformation, "Email Sent"
                            ChangeQuoteStatus
                        Else
                            MsgBox "Email Not sent.", vbInformation, "Operation Cancelled"
                        End If
                    End With
                Else
                    MsgBox "No matching email found in the Inbox Or Sent Items. Please reply To the email manually With the correct Quote Reference Number.", vbExclamation, "Email Error"
                End If
            Else
                MsgBox "Client email Not found.", vbExclamation, "Email Error"
            End If

         Exit Sub

 ErrorHandler:
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
            LogToFile "Error in EmailQuoteBtn_Click: " & Err.Description
End Sub

' Function To sanitize file names
Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer

    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i

    SanitizeFileName = fileName
End Function

' Function To log messages securely
Private Sub LogToFile(message As String)
    Dim filePath As String
    Dim fileNumber As Integer

    filePath = "E:\ent\DebugLog.txt" ' Change To a secure location
    fileNumber = FreeFile

    Open filePath For Append As #fileNumber
    Print #fileNumber, Now() & " - " & message
    Close #fileNumber
End Sub

' Function To check If the email's To field matches the recipient email
Function EmailMatchesRecipient(emailToField As String, recipientEmail As String) As Boolean
    Dim emailAddress As String
    Dim emailParts As Variant
    Dim i As Integer

    ' Split the To field into individual email addresses
    emailParts = Split(emailToField, ";")

    ' Loop through each email address in the To field
    For i = LBound(emailParts) To UBound(emailParts)
        emailAddress = Trim(emailParts(i)) ' Remove leading/trailing spaces

        ' Extract the email address If it contains a display name (e.g., "John Doe <john@example.com>")
        If InStr(emailAddress, "<") > 0 Then
            emailAddress = Mid(emailAddress, InStr(emailAddress, "<") + 1)
            emailAddress = Left(emailAddress, InStr(emailAddress, ">") - 1)
        End If

        ' Compare the normalized email address With the recipient email
        If LCase(emailAddress) = LCase(recipientEmail) Then
            EmailMatchesRecipient = True
         Exit Function
        End If
    Next i

    ' If no match is found, return False
    EmailMatchesRecipient = False
End Function

//////////////////////////////////
' sending seperate emails to primay & secondry emails 
Private Sub ChangeQuoteStatus()
    On Error Goto ErrorHandler
        Dim FatherForm As Form
        Dim db As DAO.Database
        Dim rsX As DAO.Recordset
        Dim rsY As DAO.Recordset
        Dim strSQL As String
        Dim QRequestRefNo As String ' Assuming QRequestRefNo is a string; change To Long If it's a number

        ' Initialize database And recordsets
        Set db = CurrentDb
        Set FatherForm = Forms("FuelSupport_QuoteRequestF_V1").Form

        ' Get the QuoteRequestRefNo value (replace this With your actual logic To Get the value)
        QRequestRefNo = Me.R_RefNo ' Example: Get the value from a form control

        ' Open Table Y (FuelSupport_FuelQuotationsT) With a filter For the specific QuoteRequestRefNo
        Set rsY = db.OpenRecordset("Select * FROM FuelSupport_FuelQuotationsT WHERE QuoteRequestRefNo = '" & QRequestRefNo & "'", dbOpenDynaset)

        ' Check If Table Y has records
        If rsY.EOF And rsY.BOF Then
            MsgBox "No records found in Table Y For QuoteRequestRefNo: " & QRequestRefNo, vbExclamation
         Exit Sub
        End If

        ' Loop through Table Y And update Table X For matching locations
        rsY.MoveFirst
        Do While Not rsY.EOF
            ' Open Table X With a filter For the current location And QuoteRequestRefNo
            strSQL = "Select * FROM FuelSupport_RequestedLocationsT_V1 WHERE QuoteRequestRefNo = '" & QRequestRefNo & "' And Location = '" & rsY!AIRPORT & "'"
            Set rsX = db.OpenRecordset(strSQL, dbOpenDynaset)

            ' If a matching record is found, update Table X
            If Not rsX.EOF And Not rsX.BOF Then
                rsX.MoveFirst
                Do While Not rsX.EOF
                    rsX.Edit
                    rsX!LocationQuoteStatus = "Sent"
                    rsX!QuoteRefNo = rsY!Quote_RefNo '' Update values in Table X With QuoteStatus from Table Y
                    rsX.Update
                    rsX.MoveNext
                Loop
            End If

            ' Close the Table X recordset For the current location
            rsX.Close
            Set rsX = Nothing

            ' Move To the Next record in Table Y
            rsY.MoveNext
        Loop

        ' Clean up
        rsY.Close
        Set rsY = Nothing
        Set db = Nothing

        MsgBox "Table X has been updated successfully For QuoteRequestRefNo: " & QRequestRefNo, vbInformation
     Exit Sub

        FatherForm.Controls("FuelSupport_LocationsListF").Form.Requery

 ErrorHandler:
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        ' Clean up in Case of error
        If Not rsX Is Nothing Then rsX.Close
            If Not rsY Is Nothing Then rsY.Close
                Set rsX = Nothing
                Set rsY = Nothing
                Set db = Nothing
End Sub

Private Sub EmailQuoteBtn_Click()
    On Error Goto ErrorHandler

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

        ' Apply filter To the form
        Me.Filter = strWhere
        Me.FilterOn = True

        ' Extract values from the report
        Dim AIRPORT As String
        Dim clientID As String
        Dim recipientEmail As String
        Dim secondaryEmail As String
        Dim fileName As String
        Dim rs As DAO.Recordset

        ' Use parameterized query To prevent SQL injection
        Dim qdf As DAO.QueryDef
        Set qdf = CurrentDb.CreateQueryDef("", "Select DISTINCT Airport FROM FuelSupport_FuelQuotationsT WHERE [Quote_RefNo] Like @QuoteRefNo")
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
            secondaryEmail = Nz(DLookup("SecondaryEmail", "CustomersT", "ID = " & clientID), "")
            LogToFile "Client email: " & recipientEmail
            LogToFile "Client secondary email: " & secondaryEmail

            ' Construct And sanitize file name
            fileName = SanitizeFileName(Me.Quote_RefNotxt & " " & AIRPORT & " " & clientID & ".pdf")
            LogToFile "File name: " & fileName

            ' Open file dialog To Select path
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

            ' Export report To PDF
            DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
            DoCmd.OutputTo acOutputReport, "FuelSupport_FuelQuote", acFormatPDF, filePath, False
            DoCmd.Close acReport, "FuelSupport_FuelQuote"
            LogToFile "Report exported To: " & filePath

            ' Send email If recipient is found
            If recipientEmail <> "" Then
                Dim olApp As Object
                Dim olMailItem As Object
                Dim currentUserEmail As String

                ' Initialize Outlook application
                Set olApp = CreateObject("Outlook.Application")
                Set olMailItem = olApp.CreateItem(0) ' 0 = olMailItem

                ' Get the current user's email address
                currentUserEmail = olApp.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                LogToFile "Current user email: " & currentUserEmail

                ' Set email properties
                With olMailItem
                    .SentOnBehalfOfName = "occ@tahseenaviation.com" ' Send on behalf of OCC
                    .To = recipientEmail
                    If secondaryEmail <> "" Then
                        .To = .To & ";" & secondaryEmail ' Add secondary email To the "To" field
                    End If
                    .CC = currentUserEmail & ";occ@tahseenaviation.com" ' CC current user And OCC
                    .Subject = "Tahseen Fuels Quote #" & Me.Quote_RefNotxt & " For " & AIRPORT
                    .Attachments.Add filePath
                    .HTMLBody = "Dear On Duty, " & "<br><br>" & _
                    "Thank you For your inquiry And For choosing Tahseen Aviation Services." & "<br><br>" & _
                    "Tahseen Fuels Quote #" & Me.Quote_RefNotxt & " For " & AIRPORT & " has been forwarded To your email address. We kindly request that you send us the confirmed flight details To proceed With the order." & "<br><br>" & _
                    "Please Do Not hesitate To contact us If you have any further fuel requests Or inquiries." & "<br><br>" & _
                    "<br>" & .HTMLBody

                    ' Send the email after confirmation
                    If MsgBox("Are you sure you want To send this email?", vbYesNo + vbQuestion) = vbYes Then
                        .Send
                        MsgBox "Email sent successfully With the attached quote.", vbInformation, "Email Sent"
                        LogToFile "Email sent successfully."
                        ChangeQuoteStatus
                    Else
                        MsgBox "Email Not sent.", vbInformation, "Operation Cancelled"
                        LogToFile "Email sending cancelled by user."
                    End If
                End With
            Else
                MsgBox "Client email Not found.", vbExclamation, "Email Error"
                LogToFile "Client email Not found."
            End If

         Exit Sub

 ErrorHandler:
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
            LogToFile "Error in EmailQuoteBtn_Click: " & Err.Description
End Sub

' Function To sanitize file names
Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer

    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i

    SanitizeFileName = fileName
End Function

' Function To log messages securely
Private Sub LogToFile(message As String)
    Dim filePath As String
    Dim fileNumber As Integer

    filePath = "E:\ent\DebugLog.txt" ' Change To a secure location
    fileNumber = FreeFile

    Open filePath For Append As #fileNumber
    Print #fileNumber, Now() & " - " & message
    Close #fileNumber
End Sub