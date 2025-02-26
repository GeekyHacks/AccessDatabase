'=======================================================
' Helper Function: GeneratePartyCode
'=======================================================
Private Function GeneratePartyCode() As Long
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Dim newCorpID As Long

        ' Step 1: Retrieve the last CorpID from the PartyCodeT_V13 table
        Set rst = CurrentDb.OpenRecordset("PartyCodeT_V13", dbOpenDynaset)

        If Not rst.EOF And Not IsNull(rst!corpID) Then
            newCorpID = rst!corpID + 1 ' Increment the last CorpID
        Else
            newCorpID = 1 ' Initialize To 1 If no records exist
        End If

        ' Step 2: Update the PartyCodeT_V13 table With the New CorpID
        If rst.recordCount > 0 Then
            rst.Edit ' Update the existing record
        Else
            rst.AddNew ' Add a New record If the table is empty
        End If
        rst!corpID = newCorpID
        rst.Update

        rst.Close
        Set rst = Nothing

        ' Return the New CorpID
        GeneratePartyCode = newCorpID
     Exit Function

 ErrorHandler:
        LogError "Error in GeneratePartyCode: " & Err.Description, "GeneratePartyCode"
        MsgBox "Error in GeneratePartyCode: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
        GeneratePartyCode = -1
End Function

'=======================================================
' Helper Function: AddParty
'=======================================================
Private Function AddParty() As Long
    On Error Goto ErrorHandler

        ' Validate required fields
        If IsNull(Me.PartyName.Value) Or IsNull(Me.FSANumber.Value) Then
            MsgBox "Party Name And FSANumber are required.", vbExclamation
            AddParty = -1
         Exit Function
        End If

        Dim rst As DAO.Recordset
        Dim corpID As Long
        Dim PartyCode As String

        ' Step 1: Generate the Next CorpID
        corpID = GeneratePartyCode()
        If corpID = -1 Then Exit Function ' Exit If there was an error

            ' Step 2: Generate the PartyCode dynamically
            If Me.VendorChx.Value = True Then
                PartyCode = "CV" & Format(corpID, "0000") & Format(date, "yy") ' Format As CVXXXXYY
            Else
                PartyCode = "C" & Format(corpID, "0000") & Format(date, "yy") ' Format As CXXXXYY
            End If

            ' Step 3: Add the New party To the PartiesT_V13 table
            Set rst = CurrentDb.OpenRecordset("PartiesT_V13", dbOpenDynaset)

            With rst
                .AddNew
                !corpID = corpID
                !PartyCode = PartyCode
                !PartyName = Me.PartyName.Value
                !FSANumber = Me.FSANumber.Value
                !AgreementStatus = Me.AgreementStatus.Value
                !AgreementExpiry = Me.AgreementExpiry.Value ' One year date from today
                !DateSigned = Me.DateSigned.Value
                !VAT_NO = Me.VAT_NO.Value
                ' !DocStorage = Me.DocStorage.Value ' Uncomment If needed
                .Update

                ' Retrieve the newly added CorpID
                .Bookmark = .LastModified
                AddParty = !corpID
            End With

            rst.Close
            Set rst = Nothing

            ' Display the PartyCode on the form
            DisplayPartyCode corpID

         Exit Function

 ErrorHandler:
            LogError "Error in AddParty: " & Err.Description, "AddParty"
            MsgBox "Error in AddParty: " & Err.Description, vbCritical
            If Not rst Is Nothing Then
                rst.Close
                Set rst = Nothing
            End If
            AddParty = -1
End Function

'=======================================================
'This Function generates the PartyCode dynamically For display purposes
'=======================================================
Private Function GenerateDisplayPartyCode(corpID As Long) As String
    Dim prefix As String

    ' Determine the prefix based on VendorChx
    If Me.VendorChx.Value = True Then
        prefix = "CV" ' Use "CV" For clients that are also vendors
    Else
        prefix = "C" ' Use "C" For clients
    End If

    ' Generate the PartyCode in the format [Prefix][XXXX][YY]
    GenerateDisplayPartyCode = prefix & Format(corpID, "0000") & Format(date, "yy")
End Function

'=======================================================
'Displays the dynamically generated PartyCode on the form.
'=======================================================
Private Sub DisplayPartyCode(corpID As Long)
    Me.PartyCode.Value = GenerateDisplayPartyCode(corpID)
End Sub

'=======================================================
' Main Function: AddClientBtn_Click
'=======================================================
Private Sub AddPartyBtn_Click()
    On Error Goto ErrorHandler

        ' Add Party
        Dim partyID As Long
        partyID = AddParty()
        If partyID = -1 Then Exit Sub ' Exit If there was an error

            ' Add Location
            Dim locationID As Long
            locationID = AddLocation(partyID)
            If locationID = -1 Then Exit Sub

                ' Add Party Contact
                AddPartyContact partyID, locationID

                ' Add Client
                AddClient partyID

                ' Add Roles (from ListBox)
                Dim roleID As Long
                Dim i As Variant
                For Each i In Me.RolesListBox.ItemsSelected
                    roleID = Me.RolesListBox.Column(0, i) ' Assuming the Role ID is in the first column
                    AddPartyRole partyID, roleID
                Next i

                ' Add Vendor (If applicable)
                If Me.VendorChx.Value = True Then
                    ' Add Vendor
                    AddVendor partyID

                    ' Add Party Location (Vendor)
                    AddPartyLocation partyID, locationID
                End If

                MsgBox "Client added successfully!", vbInformation
             Exit Sub

 ErrorHandler:
                LogError "Error in AddPartyBtn_Click: " & Err.Description, "AddPartyBtn_Click"
                MsgBox "Error in AddPartyBtn_Click: " & Err.Description, vbCritical
End Sub

'=======================================================
' Helper Function: AddPartyRole
'=======================================================
Private Sub AddPartyRole(partyID As Long, roleID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("PartyRolesJT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !corpID = partyID ' Use the passed CorpID
            !roleID = roleID
            !IsActive = True
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddPartyRole: " & Err.Description, "AddPartyRole"
        MsgBox "Error in AddPartyRole: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub

'=======================================================
' Helper Function: AddLocation
'=======================================================
Private Function AddLocation(partyID As Long) As Long
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("LocationsT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !corpID = partyID ' Use the passed CorpID
            !Country = Me.Country.Value
            !City = Me.City.Value
            !Address = Me.Address.Value
            .Update

            ' Retrieve the newly added Location ID
            .Bookmark = .LastModified
            AddLocation = !ID
        End With

        rst.Close
        Set rst = Nothing
     Exit Function

 ErrorHandler:
        LogError "Error in AddLocation: " & Err.Description, "AddLocation"
        MsgBox "Error in AddLocation: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
        AddLocation = -1
End Function

'=======================================================
' Helper Function: AddPartyContact
'=======================================================
Private Sub AddPartyContact(partyID As Long, locationID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("PartyContactT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !corpID = partyID
            !locationID = locationID
            !PrimaryEmail = Me.PrimaryEmail.Value
            !SecondaryEmail = Me.SecondaryEmail.Value
            !Phone = Me.Phone.Value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddPartyContact: " & Err.Description, "AddPartyContact"
        MsgBox "Error in AddPartyContact: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub

'=======================================================
' Helper Function: AddClient
'=======================================================
Private Sub AddClient(partyID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("ClientsT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !corpID = partyID
            !CreditLimit = Me.CreditLimit.Value
            !Deposit = Me.Deposit.Value
            !Currency = Me.Currency.Value
            !ClientRating = Me.ClientRating.Value
            !ClientDescription = Me.ClientDescription.Value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddClient: " & Err.Description, "AddClient"
        MsgBox "Error in AddClient: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub

'=======================================================
' Helper Function: AddVendor
'=======================================================
Private Sub AddVendor(partyID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("VendorsT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !corpID = partyID
            !CreditLimit = Me.CreditLimit.Value
            !Deposit = Me.Deposit.Value
            !Currency = Me.Currency.Value
            !VendorRating = Me.VendorRating.Value
            !VendorDescription = Me.ClientDescription.Value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddVendor: " & Err.Description, "AddVendor"
        MsgBox "Error in AddVendor: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub

'=======================================================
' Helper Function: AddPartyLocation
'=======================================================
Private Sub AddPartyLocation(partyID As Long, locationID As Long)
    On Error Goto ErrorHandler

        Dim rst As DAO.Recordset
        Set rst = CurrentDb.OpenRecordset("PartyLocations_JT_V13", dbOpenDynaset)

        With rst
            .AddNew
            !corpID = partyID
            !locationID = locationID
            !IsPrimary = True
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        LogError "Error in AddPartyLocation: " & Err.Description, "AddPartyLocation"
        MsgBox "Error in AddPartyLocation: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub

'=======================================================
' LogError Subroutine: Logs errors To a text file
'=======================================================
Private Sub LogError(ErrorMessage As String, ErrorSource As String)
    On Error Resume Next ' Prevent infinite Loop If logging fails

    Dim filePath As String
    Dim fileNumber As Integer
    Dim logMessage As String

    ' Define the file path
    filePath = "E:\ent\DebugLog_AddClientF_V13.txt"

    ' Open the file For appending (creates the file If it doesn't exist)
    fileNumber = FreeFile
    Open filePath For Append As #fileNumber

    ' Prepare the log message
    logMessage = "Error Date: " & Now() & vbCrLf & _
    "Error Source: " & ErrorSource & vbCrLf & _
    "Error Message: " & ErrorMessage & vbCrLf & _
    "----------------------------------------" & vbCrLf

    ' Write the log message To the file
    Print #fileNumber, logMessage

    ' Close the file
    Close #fileNumber
End Sub
'=======================================================
' Subroutine: ResetForm
' Purpose: Clears all input fields To prepare For a New record.
'=======================================================
Private Sub ResetForm()
    On Error Resume Next ' Skip errors For controls without a Value Property

    ' Clear text boxes, combo boxes, And other input fields
    Me.PartyCode.Value = Null
    Me.PartyName.Value = Null
    Me.FSANumber.Value = Null
    Me.AgreementStatus.Value = "Pending" ' "Pending" is the default value
    Me.AgreementExpiry.Value = Null
    Me.DateSigned.Value = Null
    Me.VAT_NO.Value = Null
    Me.Country.Value = Null
    Me.City.Value = Null
    Me.Address.Value = Null
    Me.PrimaryEmail.Value = Null
    Me.SecondaryEmail.Value = Null
    Me.Phone.Value = Null
    Me.CreditLimit.Value = 15000 ' 15000 is the default value
    Me.Deposit.Value = Null
    Me.Currency.Value = "USD"  ' "USD" is the default value
    Me.ClientRating.Value = 3 ' 3 is the default value
    Me.ClientDescription.Value = Null
    Me.VendorRating.Value = 3 ' 3 is the default value


    ' Reset checkboxes And toggle buttons To default state
    Me.VendorChx.Value = False
    Me.IsActive.Value = True ' Default To "Active"

    ' Set focus To the first input field
    Me.PartyName.SetFocus

    On Error Goto 0 ' Reset error handling
End Sub
Private Sub ResetBtn_Click()
    ' Ask the user For confirmation before resetting the form
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want To reset the form? All unsaved data will be lost.", vbYesNo + vbQuestion, "Confirm Reset")

    ' Check the user's response
    If response = vbYes Then
        ResetForm ' Reset the form If the user clicks Yes
        LogError "Form reset by user.", "ResetBtn_Click" ' Optional logging
    Else
        ' Do nothing If the user clicks No
     Exit Sub
    End If
End Sub

above code is working good; problem is; data is added To tables As long As the code run; but If lets say error occur in AddLocation; code moves nxt regardless ??
And data is added To every table related To a working Function.
I want the whole data adding To stop in Case there's an error in any of the functions inside AddClientBtn_Click