Private Sub EmailQuoteBtn_Click()
    On Error Goto ErrorHandler

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
                MsgBox "No parameter specified. Generating report For all data...", vbCritical, "My Flight App"
            Else
                strWhere = Left$(strWhere, InLeng)

                ' Apply the filter To the form
                Me.Filter = strWhere
                Me.FilterOn = True

                ' Extract values from the report
                Dim AIRPORT As String
                Dim clientID As String
                Dim recipientEmail As String
                Dim fileName As String
                Dim rs As DAO.Recordset

                Set rs = CurrentDb.OpenRecordset("Select DISTINCT Airport FROM FuelSupport_FuelQuotationsT WHERE " & strWhere, dbOpenSnapshot)

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

                    ' Construct the file name And sanitize it
                    fileName = SanitizeFileName(QRefNo & " " & AIRPORT & " " & clientID & ".pdf")

                    ' Debugging: Show the file name
                    LogToFile "File Name: " & fileName

                    ' Open a file dialog To Select the path
                    Dim fd As fileDialog
                    Set fd = Application.fileDialog(msoFileDialogFolderPicker)
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

                    ' Apply the filter To the report before exporting
                    DoCmd.OpenReport "FuelSupport_FuelQuote", acViewPreview, , strWhere
                    DoCmd.OutputTo acOutputReport, "FuelSupport_FuelQuote", acFormatPDF, filePath, False
                    DoCmd.Close acReport, "FuelSupport_FuelQuote"

                    ' Debugging: Confirm report export
                    LogToFile "Report exported successfully To: " & filePath

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
                        LogToFile "Searching For emails With subject containing: " & searchSubject

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

                        ' Loop through emails in the Inbox To find the specific email
                        For Each olMailItem In olFolder.Items
                            ' Check If the item is a mail item
                            If olMailItem.Class = 43 Then ' 43 = olMail
                                ' Check For the specific email based on subject
                                If InStr(olMailItem.Subject, searchSubject) > 0 Then
                                    ' Reply-all To the email
                                    Set olReply = olMailItem.ReplyAll
                                    With olReply
                                        ' .SentOnBehalfOfName = "occ@tahseenaviation.com"
                                        .CC = "Abdullah@tahseenaviation.com" ' Replace With your email
                                        .Attachments.Add filePath
                                        .Body = "Dear On Duty, " & vbCrLf & vbCrLf & _
                                        "Thank you For your inquiry And For choosing Tahseen Aviation Services." & vbCrLf & vbCrLf & _
                                        "Tahseen Fuels Quote #" & QRefNo & " has been forwarded To your email address. We kindly request that you send us the confirmed flight details To proceed With the order." & vbCrLf & vbCrLf & _
                                        "Please Do Not hesitate To contact us If you have any further fuel requests Or inquiries." & vbCrLf & vbCrLf & _
                                        "Kind Regards," & vbCrLf & .Body
                                        .Send
                                    End With

                                    foundEmail = True
                                 Exit For
                                End If
                            End If
                        Next olMailItem

                        If foundEmail Then
                            MsgBox "Reply-all email sent With the attached quote.", vbInformation, "Email Sent"
                            LogToFile "Reply-all email sent To: " & recipientEmail & " With CC To occ@tahseenaviation.com And your_email@domain.com"
                        Else
                            MsgBox "No matching email found in the Inbox. Please reply To the email manually With the correct R_RefNo.", vbExclamation, "Email Error"
                            LogToFile "No matching email found in the Inbox For R_RefNo: " & searchSubject
                        End If
                    Else
                        MsgBox "Client email Not found.", vbExclamation, "Email Error"
                        LogToFile "Client email Not found For ClientID: " & clientID
                    End If

                    ' The file is now saved permanently at the selected file path
                    ' No need To delete it using Kill filePath
                End If
            Else
                DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
            End If

         Exit Sub

 ErrorHandler:
            If Err.Number = 2501 Then
                MsgBox "The OutputTo action was canceled. Please ensure the report is properly generated And the file path is correct.", vbExclamation, "Error"
                LogToFile "Error 2501: The OutputTo action was canceled."
            Else
                MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
                LogToFile "Error " & Err.Number & ": " & Err.Description
            End If
End Sub

' Function To sanitize file names by replacing invalid characters With underscores
Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer

    ' List of invalid characters in file names
    invalidChars = "\/:*?""<>|"

    ' Replace each invalid character With an underscore
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i

    ' Return the sanitized file name
    SanitizeFileName = fileName
End Function

Private Sub LogToFile(message As String)
    Dim filePath As String
    Dim fileNumber As Integer

    ' Specify the path To the log file
    filePath = "E:\ent\DebugLog.txt" ' Change this To your desired path

    ' Get the Next available file number
    fileNumber = FreeFile

    ' Open the file For appending
    Open filePath For Append As #fileNumber

    ' Write the message To the file
    Print #fileNumber, Now() & " - " & message

    ' Close the file
    Close #fileNumber
End Sub

