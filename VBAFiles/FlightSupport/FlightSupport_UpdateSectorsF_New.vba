Option Compare Database
Private Sub Form_Open(Cancel As Integer)
    Me.FilterOn = False
End Sub

Private Sub ACRegCb_Change()
    Dim selectedValue As Variant
    selectedValue = Me.ACRegCb.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("ACRegT", dbOpenSnapshot)

        rs.FindFirst "ACReg = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ACType.Value = rs("ACType").Value
            Me.MTOW.Value = rs("ACMTOW").Value
            Me.Operator.Value = rs("Operator").Value
            Me.Customer.Value = rs("Customer").Value
        Else
            Me.ACType.Value = Null
            Me.MTOW.Value = Null
            Me.Operator.Value = Null
            Me.Customer.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ACType.Value = Null
        Me.MTOW.Value = Null
        Me.Operator.Value = Null
        Me.Customer.Value = Null
    End If

 Exit Sub

End Sub

Private Sub UpdateUser1()
    Dim UserID As Integer
    UserID = Forms!loggedInUser!txtUserID
    'UserID = 1
    Me.User1 = DLookup("UserName", "UsersT", "ID =" & UserID & "")
End Sub

Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "FlightSupport_AddSectorsF", acSaveNo
End Sub
Public Sub DepartDate_Click()
    InputDateField DepartDate, "Select a date To use this on your form"
End Sub
Public Sub ArrivalDate_Click()
    InputDateField ArrivalDate, "Select a date To use this on your form"
End Sub

Private Sub TripStatuscmb_AfterUpdate()
    change_PermitServices
    change_LocationServices

End Sub
Private Sub change_PermitServices()

    On Error Goto ErrorHandler
        Dim strSQL As String
        Dim rs As DAO.Recordset
        Dim db As DAO.Database
        Dim sectID As Long

        Dim newStatus As String

        ' Check If TripStatuscmb is Not null
        If Not IsNull(Me.TripStatuscmb) Then

            ' Determine the New status based on TripStatuscmb
            Select Case Me.TripStatuscmb
             Case "CANCELED"
                newStatus = "CANCELED"
             Case "ACTIVE"
                newStatus = "CONFIRMED"
             Case Else
                MsgBox "Error: Invalid TripStatus.", vbCritical
             Exit Sub
            End Select

            sectID = Me.Sect_ID
            If IsNull(sectID) Or Not IsNumeric(sectID) Then
                MsgBox "Error: Sect_ID is invalid.", vbCritical
             Exit Sub
            End If

            Set db = CurrentDb

            ' Update PermitBriefT_New
            Set rs = db.OpenRecordset("Select Sect_ID FROM PermitBriefT_New WHERE Sect_ID = " & sectID)
            Do While Not rs.EOF
                ' Debugging message
                Debug.Print "Sect_ID: " & sectID

                strSQL = "UPDATE PermitBriefT_New Set PermitStatus = '" & newStatus & "' WHERE Sect_ID = " & sectID
                Debug.Print "Executing SQL: " & strSQL
                db.Execute strSQL
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
            Set db = Nothing
        End If

 Cleanup:
        On Error Resume Next
        If Not rs Is Nothing Then rs.Close
            Set rs = Nothing
            Set db = Nothing
         Exit Sub

 ErrorHandler:
            MsgBox "Error: " & err.Description, vbCritical
            Resume Cleanup

End Sub
Private Sub change_LocationServices()

    On Error Goto ErrorHandler
        Dim strSQL As String
        Dim rs As DAO.Recordset
        Dim db As DAO.Database
        Dim sectID As Long
        Dim locID As Long
        Dim newStatus As String

        ' Check If TripStatuscmb is Not null
        If Not IsNull(Me.TripStatuscmb) Then

            ' Determine the New status based on TripStatuscmb
            Select Case Me.TripStatuscmb
             Case "CANCELED"
                newStatus = "CANCELED"
             Case "ACTIVE"
                newStatus = "CONFIRMED"
             Case Else
                MsgBox "Error: Invalid TripStatus.", vbCritical
             Exit Sub
            End Select

            sectID = Me.Sect_ID
            If IsNull(sectID) Or Not IsNumeric(sectID) Then
                MsgBox "Error: Sect_ID is invalid.", vbCritical
             Exit Sub
            End If

            Set db = CurrentDb
            ' Open recordset For LocationsT_New
            ' Use dbOpenDynaset For a dynamic recordset (allows updates)
            Set rs = db.OpenRecordset("Select ID FROM LocationsT_New WHERE Sect_ID = " & sectID)

            ' Loop through each record in LocationsT_New where Sect_ID matches
            Do While Not rs.EOF

                ' Ensure Location_ID is valid
                If IsNull(rs!ID) Or Not IsNumeric(rs!ID) Then
                    MsgBox "Error: Location_ID is invalid For Sect_ID: " & sectID, vbCritical
                 Exit Sub
                End If

                locID = CLng(rs!ID)

                ' Debugging message
                Debug.Print "Location_ID: " & locID

                ' Update HandlingT_New
                strSQL = "UPDATE HandlingT_New Set ServiceStatus = '" & newStatus & "' WHERE Location_ID = " & locID
                Debug.Print "Executing SQL: " & strSQL
                db.Execute strSQL

                ' Update FuelBriefT_New
                strSQL = "UPDATE FuelBriefT_New Set FuelStatus = '" & newStatus & "' WHERE Location_ID = " & locID
                Debug.Print "Executing SQL: " & strSQL
                db.Execute strSQL

                ' Update ConciergeBriefT_New
                strSQL = "UPDATE ConciergeBriefT_New Set ServiceStatus = '" & newStatus & "' WHERE Location_ID = " & locID
                Debug.Print "Executing SQL: " & strSQL
                db.Execute strSQL

                rs.MoveNext
            Loop
            rs.Close
        End If

        Me.Requery

 Cleanup:
        On Error Resume Next
        If Not rs Is Nothing Then rs.Close
            Set rs = Nothing
            Set db = Nothing
         Exit Sub

 ErrorHandler:
            MsgBox "Error: " & err.Description, vbCritical
            Resume Cleanup

End Sub
