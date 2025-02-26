Option Compare Database
Option Explicit
Private Sub City_GotFocus()
    If IsNull(Me.Country.Value) Then
        MsgBox "Select a country first"
        Me.Country.SetFocus
    End If
End Sub

Private Sub CloseBtn_Click()
    'ResetFormFields
    Me.Undo
    DoCmd.Close acForm, "AddClientF_V13"
End Sub
Private Sub Country_AfterUpdate()
    Dim rst As DAO.Recordset
    Dim Nationality As DAO.Recordset
    Dim selectedCountry As Variant
    Dim nationalityValue As String
    Dim sql As String

    On Error Goto ErrorHandler

        ' Get the selected country code
        selectedCountry = Me.Country.Column(0)

        ' Sanitize And build the SQL query For nationality
        sql = SanitizeSQL("Select nationality FROM CountriesT_V13 WHERE num_code = [param1];", selectedCountry)
        Set Nationality = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

        ' Check If the recordset is Not empty
        If Not Nationality.EOF Then
            nationalityValue = Nationality!Nationality
        Else
            nationalityValue = "" ' Or some default value
        End If

        ' Set the nationality value in the form
        'Me.Nationality.Value = nationalityValue

        ' Close the nationality recordset
        Nationality.Close
        Set Nationality = Nothing

        ' Sanitize And build the SQL query For cities
        sql = SanitizeSQL("Select City FROM CitiesT_V13 WHERE CountryCode = [param1];", selectedCountry)
        Me.City.RowSource = sql
        Me.City.Requery

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & Err.Description
        If Not Nationality Is Nothing Then
            Nationality.Close
            Set Nationality = Nothing
        End If
End Sub
' Reusable Function To sanitize SQL queries With parameters
Function SanitizeSQL(sql As String, ParamArray params() As Variant) As String
    Dim i As Integer
    Dim paramValue As Variant

    ' Replace each placeholder With the sanitized parameter value
    For i = LBound(params) To UBound(params)
        paramValue = params(i)

        ' Sanitize the parameter value (e.g., escape single quotes)
        If IsNull(paramValue) Then
            paramValue = "NULL"
        Elseif IsNumeric(paramValue) Then
            ' Numeric values don't need escaping
        Else
            ' Escape single quotes For strings
            paramValue = "'" & Replace(paramValue, "'", "''") & "'"
        End If

        ' Replace the placeholder With the sanitized value
        sql = Replace(sql, "[param" & (i + 1) & "]", paramValue)
    Next i

    ' Return the sanitized SQL query
    SanitizeSQL = sql
End Function
Private Sub DateSigned_Click()
    InputDateField DateSigned, "Select a date To use this on your form"
End Sub
Private Sub AOCExpiry_Click()
    InputDateField AOCExpiry, "Select a date To use this on your form"
End Sub
Public Sub ResetBtn_Click()
    On Error Goto ErrorHandler

        Me.Undo
        DoCmd.CancelEvent
        DoCmd.GoToRecord , , acNewRec

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & Err.Description, vbExclamation
End Sub

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
                !IsActive = Me.IsActive.Value
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
            locationID = AddLocation()
            If locationID = -1 Then Exit Sub

                ' Add Party Contact
                AddPartyContact partyID, locationID

                ' Add Client
                AddClient partyID

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
                MsgBox "Error in AddClientBtn_Click: " & Err.Description, vbCritical
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
            !partyID = partyID ' Use the passed CorpID
            !roleID = roleID
            !IsActive = True
            !CreatedDate = date
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
        MsgBox "Error in AddPartyRole: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub
'=======================================================
' Helper Function: AddLocation
'=======================================================
Private Function AddLocation() As Long
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
            !IsActive = Me.IsActive.Value
            !ClientDescription = Me.ClientDescription.Value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
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
            !IsActive = Me.IsActive.Value
            !VendorDescription = Me.ClientDescription.Value
            .Update
        End With

        rst.Close
        Set rst = Nothing
     Exit Sub

 ErrorHandler:
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
        MsgBox "Error in AddPartyLocation: " & Err.Description, vbCritical
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
End Sub
Private Sub ResetFormFields()
    ' Clear all form fields
    Me.Code.Value = Null
    Me.OperatorName.Value = Null
    Me.ICAO.Value = Null
    Me.IATA.Value = Null
    Me.AOCExpiry.Value = Null
    Me.Address.Value = Null
    Me.ContactInfo.Value = Null
    Me.Country.Value = Null
    Me.City.Value = Null
    Me.PrimaryEmail.Value = Null
    Me.FSANumber.Value = Null
    Me.AgreementStatus.Value = "PENDING"
    Me.DateSigned.Value = Null
    Me.SecondaryEmail.Value = Null
    Me.CargoCmb.Value = Null
    Me.PVTCmb.Value = Null
    Me.AmbulanceCmb.Value = Null
    Me.COMCmb.Value = Null
    Me.Nationality.Value = Null
    Me.AllOpsCmb.Value = Null

    ' Set focus To the first field For data entry
    Me.OperatorName.SetFocus
End Sub
Private Sub Command192_Click()
    DoCmd.OpenForm "UpdateOperatorF", WindowMode:=acWindowNormal
    CloseBtn_Click
End Sub
Private Sub Form_Load()
    Dim fw As New clFormWindow
    DoCmd.GoToRecord , , acNewRec
    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
    Me.VendorRating.Visible = False

End Sub
Private Sub VendorChx_Click()
    If Me.VendorChx.Value = -1 Then
        Me.VendorRating.Visible = True
    Else
        Me.VendorRating.Visible = False
    End If
End Sub
