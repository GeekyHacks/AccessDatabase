Private Sub EmailQuoteBtn_Click()
    On Error GoTo ErrorHandler

    If Not IsNull(Me.Quote_RefNotxt) Then
        Dim strWhere As String
        Dim InLeng As Long

        Const conJetDate = "\#dd\/mm\/yyyy\#"

        If Not IsNull(Me.Quote_RefNotxt) Then
            strWhere = strWhere & "([Quote_RefNo] Like '*" & Me.Quote_RefNotxt & "*') And "
        End If
        Dim QRefNo As Variant
        QRefNo = Me.Quote_RefNotxt.Value

        InLeng = Len(strWhere) - 5
        If InLeng <= 0 Then
            MsgBox "No parameter specified. Generating report for all data...", vbCritical, "My Flight App"
        Else
            strWhere = Left$(strWhere, InLeng)

            ' Apply the filter to the form
            Me.Filter = strWhere
            Me.FilterOn = True

            ' Extract values from the report
            Dim AIRPORT As String
            Dim clientID As String
            Dim recipientEmail As String
            Dim fileName As String
            Dim rs As DAO.Recordset

            Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT Airport FROM FuelSupport_FuelQuotationsT WHERE " & strWhere, dbOpenSnapshot)

            AIRPORT = ""
            Do While Not rs.EOF
                ' Extract the first four characters from the airport name
                AIRPORT = AIRPORT & Left(rs!AIRPORT, 4) & "_"
                rs.MoveNext
            Loop

            ' Remove trailing underscores
            If Len(AIRPORT) > 0 Then AIRPORT = Left(AIRPORT, Len(AIRPORT) - 1)

            rs.Close
            Set rs = Nothing

            clientID = Me.ClientIDtxt.Value
            recipientEmail = Nz(DLookup("PrimaryEmail", "CustomersT", "ID = " & clientID), "")

            ' Construct the file name and sanitize it
            fileName = SanitizeFileName(QRefNo & " " & AIRPORT & " " & clientID & ".pdf")

            ' Debugging: Show the file name
            LogToFile "File Name: " & fileName

            ' Open a file dialog to select the path
            Dim fd As FileDialog
            Set fd = Application.FileDialog(msoFileDialogFolderPicker)
            Dim filePath As String

            With fd
                .Title = "Select Folder"
                .AllowMultiSelect = False
                If .Show = -1 Then
                    filePath = .SelectedItems(1) & "\" & fileName
                    ' Debugging: Show the selected file path
                    LogToFile "Selected File Path: " & filePath
                Else
                    MsgBox "No folder selected. Operation cancelled.", vbExclamation
                    Exit Sub
                End If
            End With

            ' Apply the filter to the report before exporting
            DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
            DoCmd.OutputTo acOutputReport, "FuelSupport_FuelQuote", acFormatPDF, filePath, False
            DoCmd.Close acReport, "FuelSupport_FuelQuote"

            ' Debugging: Confirm report export
            LogToFile "Report exported successfully to: " & filePath

            If recipientEmail <> "" Then
                ' Find the specific email in Outlook based on searchSubject
                Dim olApp As Object
                Dim olNamespace As Object
                Dim olFolder As Object
                Dim olMailItem As Object
                Dim olReply As Object
                Dim foundEmail As Boolean
                Dim searchSubject As String

                ' Construct the search subject (e.g., "12345")
                searchSubject = CStr(Me.R_RefNo)
                LogToFile "Searching for emails with subject containing: " & searchSubject

                Set olApp = CreateObject("Outlook.Application")
                Set olNamespace = olApp.GetNamespace("MAPI")
                Set olFolder = olNamespace.GetDefaultFolder(6) ' 6 = Inbox folder

                ' Log the folder being searched
                LogToFile "Searching in folder: " & olFolder.Name

                ' Log the current user's email
                Dim currentUserEmail As String
                currentUserEmail = olNamespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                LogToFile "Current user email: " & currentUserEmail

                ' Log the subject of the last email in the inbox
                If olFolder.Items.Count > 0 Then
                    Dim lastEmail As Object
                    Set lastEmail = olFolder.Items(olFolder.Items.Count)
                    LogToFile "Last email subject in inbox: " & lastEmail.Subject & ", Sender: " & lastEmail.SenderEmailAddress & ", Receiver: " & lastEmail.To
                Else
                    LogToFile "Inbox is empty."
                End If

                foundEmail = False

                ' Loop through emails in the Inbox to find the specific email
                For Each olMailItem In olFolder.Items
                    ' Check if the item is a mail item
                    If olMailItem.Class = 43 Then ' 43 = olMail
                        ' Check for the specific email based on subject
                        If InStr(olMailItem.Subject, searchSubject) > 0 Then
                            ' Reply-all to the email
                            Set olReply = olMailItem.ReplyAll
                            With olReply
                                ' .SentOnBehalfOfName = "occ@tahseenaviation.com"
                                .CC = "Abdullah@tahseenaviation.com" ' Replace with your email
                                .Attachments.Add filePath
                                .HTMLBody = "Dear On Duty, " & "<br><br>" & _
                                            "Thank you for your inquiry and for choosing Tahseen Aviation Services." & "<br><br>" & _
                                            "Tahseen Fuels Quote #" & QRefNo & " has been forwarded to your email address. We kindly request that you send us the confirmed flight details to proceed with the order." & "<br><br>" & _
                                            "Please do not hesitate to contact us if you have any further fuel requests or inquiries." & "<br><br>" & _
                                            "Kind Regards," & "<br>" & .HTMLBody
                                .Send
                            End With

                            foundEmail = True
                            Exit For
                        End If
                    End If
                Next olMailItem

                If foundEmail Then
                    MsgBox "Reply-all email sent with the attached quote.", vbInformation, "Email Sent"
                    LogToFile "Reply-all email sent to: " & recipientEmail & " with CC to occ@tahseenaviation.com and your_email@domain.com"
                Else
                    MsgBox "No matching email found in the Inbox. Please reply to the email manually with the correct R_RefNo.", vbExclamation, "Email Error"
                    LogToFile "No matching email found in the Inbox for R_RefNo: " & searchSubject
                End If
            Else
                MsgBox "Client email not found.", vbExclamation, "Email Error"
                LogToFile "Client email not found for ClientID: " & clientID
            End If

            ' The file is now saved permanently at the selected file path
            ' No need to delete it using Kill filePath
        End If
    Else
        DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
    End If

    Exit Sub

ErrorHandler:
    If Err.Number = 2501 Then
        MsgBox "The OutputTo action was canceled. Please ensure the report is properly generated and the file path is correct.", vbExclamation, "Error"
        LogToFile "Error 2501: The OutputTo action was canceled."
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        LogToFile "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

' Function to sanitize file names by replacing invalid characters with underscores
Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer

    ' List of invalid characters in file names
    invalidChars = "\/:*?""<>|"

    ' Replace each invalid character with an underscore
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i

    ' Return the sanitized file name
    SanitizeFileName = fileName
End Function

Private Sub LogToFile(message As String)
    Dim filePath As String
    Dim fileNumber As Integer

    ' Specify the path to the log file
    filePath = "E:\ent\DebugLog.txt" ' Change this to your desired path

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the file for appending
    Open filePath For Append As #fileNumber

    ' Write the message to the file
    Print #fileNumber, Now() & " - " & message

    ' Close the file
    Close #fileNumber
End Sub