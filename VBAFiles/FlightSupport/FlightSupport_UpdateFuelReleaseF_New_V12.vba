Option Compare Database
Option Explicit
Private Sub AddFuelRelease_Click()
    On Error Goto ErrorHandler
        Dim db As DAO.Database
        Dim rs As DAO.Recordset

        If IsNull(Me.Location) Or IsNull(Me.NextDestCmb) Or IsNull(Me.ReleaseETD) Or IsNull(Me.ReleaseETA) Or IsNull(Me.ArrivalDate) Or IsNull(Me.DepartDate) Then

            MsgBox "Please add missing Fields"
         Exit Sub

        Else
            Set db = CurrentDb
            Set rs = db.OpenRecordset("FuelReleaseT_New_V12", dbOpenDynaset)

            rs.FindFirst "Location_ID = " & Me.Location.Column(1)

            If Not rs.NoMatch Then
                MsgBox "A record With this Location already exists.", vbExclamation + vbOKOnly, "Record Already Exists"
             Exit Sub
            Else

                rs.AddNew

                ' Assign values To the fields in the table
                rs.Fields("Location_ID").Value = Me.Location.Column(1)
                rs.Fields("NextDest").Value = Me.NextDestCmb.Value
                rs.Fields("ReleaseETD").Value = Me.ReleaseETD.Value
                rs.Fields("ReleaseETA").Value = Me.ReleaseETA.Value
                rs.Fields("ArrivalDate").Value = Me.ArrivalDate.Value
                rs.Fields("DepartDate").Value = Me.DepartDate.Value
                ' Save the New record
                rs.Update

                ' Clean up
                rs.Close
                Set rs = Nothing
                Set db = Nothing

                ' Reset form values
                Reset
                ' Display a message indicating successful addition
                MsgBox "Release added successfully, Please send it To client And their handler If needed"
                Me.Undo
             Exit Sub
            End If

        End If
 ErrorHandler:
        MsgBox "An unexpected error occurred: " & err.Description, vbCritical, "Error"
        Me.Undo
     Exit Sub
End Sub
Private Sub ArrivalDate_Click()
    'InputDateField ArrivalDate, "Select a date To use this on your form"
End Sub
Private Sub DepartDate_Click()
    'InputDateField DepartDate, "Select a date To use this on your form"
End Sub
Private Sub PrintRelease_Click()
    Dim strWhere As String
    Dim IngLen As Long
    Dim Loc_ID As Integer

    Loc_ID = Me.Location.Column(1)
    ' Initialize strWhere
    strWhere = ""

    If Not IsNull(Loc_ID) Then
        ' Build the WHERE clause
        strWhere = strWhere & "([Location_ID] = " & Loc_ID & ")"
    End If

    ' Check If there's a WHERE clause
    IngLen = Len(strWhere)
    If IngLen > 0 Then
        ' Apply the filter And open the report
        Me.Filter = strWhere
        Me.FilterOn = True
        DoCmd.OpenReport "FlightSupport_FuelReleaseWithoutVendor_New", acViewPreview, , strWhere
    Else
        ' No filter, open the report generator form
        DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acWindowNormal
    End If

End Sub
Private Sub Reset()
    Me.NextDestCmb.Value = Null
    Me.ReleaseETD.Value = Null
    Me.ReleaseETA.Value = Null
    Me.ArrivalDate.Value = Null
    Me.DepartDate.Value = Null
End Sub
Private Sub Location_AfterUpdate()
    TestParameterizedQuery
    RenderRecord
End Sub
Private Sub TestParameterizedQuery()
    Dim SelectedSect_ID As Integer
    Dim FlightRefNo As Variant

    SelectedSect_ID = Me.Location.Column(1)

    ' Use DLookup To directly retrieve the FlightRefNo value
    FlightRefNo = DLookup("FlightRefNo", "FuelSupport_ReleaseQ_New_V12", "Loc_ID = " & SelectedSect_ID)

    If Not IsNull(FlightRefNo) Then
        Me!RefNo.Value = FlightRefNo
    Else
        MsgBox "No FlightRefNo found For Loc_ID " & SelectedSect_ID
    End If
End Sub
Private Sub RenderRecord()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim RefNo As Variant
    Dim isRecordFound As Boolean
    Dim SelectedSectorNo As Integer
    Dim SelectedLocation As Integer

    RefNo = Me.RefNo.Value
    SelectedLocation = Me.Location.Column(1)

    ' Check If the selected location ID is Not null
    If Not IsNull(SelectedLocation) Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset("Select * FROM FuelSupport_ReleaseQ_New_V12 WHERE FlightRefNo = '" & RefNo & "'", dbOpenDynaset)

        ' Find the first record With the selected FlightRefNo
        rs.FindFirst "Loc_ID = " & SelectedLocation
        If Not rs.NoMatch Then
            If SelectedLocation = rs("Dest_Location_ID").Value Then
                SelectedSectorNo = rs("SectorNo").Value
                Me.ReleaseETA.Value = rs("Dest_ETA").Value
                Me.ArrivalDate.Value = rs("Dest_ArrivalDate").Value

                If SelectedSectorNo > 0 Then
                    rs.MoveNext
                    If rs.EOF Then
                        'rs.Close
                        MsgBox "No sector after the current sector"
                        Me.ReleaseETD.Value = "00:00"
                        Me.DepartDate.Value = "00/00/000"
                        Me.NextDestCmb.Value = "......"
                    Elseif Not rs.EOF And SelectedLocation = rs("Origin_Location_ID").Value Then
                        SelectedSectorNo = rs("SectorNo").Value
                        MsgBox "ya yao"
                        Me.ReleaseETD.Value = rs("Origin_ETD").Value
                        Me.DepartDate.Value = rs("Origin_DepartDate").Value
                        Me.NextDestCmb.Value = rs("DestAirportCode").Value
                    End If
                End If
            Elseif SelectedLocation = rs("Origin_Location_ID").Value Then
                SelectedSectorNo = rs("SectorNo").Value
                If SelectedSectorNo = 1 Then
                    Me.ReleaseETD.Value = rs("Origin_ETD").Value
                    Me.DepartDate.Value = rs("Origin_DepartDate").Value
                    Me.ReleaseETA.Value = "00:01"
                    Me.ArrivalDate.Value = rs("Origin_DepartDate").Value
                    rs.MoveNext
                    If Not rs.EOF Then
                        Me.NextDestCmb.Value = rs("DestAirportCode").Value
                    End If
                Else
                    Me.ReleaseETD.Value = rs("Origin_ETD").Value
                    Me.DepartDate.Value = rs("Origin_DepartDate").Value
                    Me.NextDestCmb.Value = rs("DestAirportCode").Value
                    SelectedSectorNo = rs("SectorNo").Value
                    If SelectedSectorNo > 1 Then
                        rs.MovePrevious
                        If Not rs.BOF Then
                            SelectedSectorNo = rs("SectorNo").Value
                            Me.ReleaseETA.Value = rs("Dest_ETA").Value
                            Me.ArrivalDate.Value = rs("Dest_ArrivalDate").Value
                            rs.MoveNext
                            rs.MoveNext
                            Me.NextDestCmb.Value = rs("DestAirportCode").Value
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "No matching records found For the selected FlightRefNo.", vbExclamation + vbOKOnly, "Record Not Found"
        End If
    Else
        MsgBox "Selected location ID is null.", vbExclamation + vbOKOnly, "Error"
    End If

    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        Set db = Nothing
    End If
End Sub

