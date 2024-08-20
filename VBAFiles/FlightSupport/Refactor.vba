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
                    Elseif Not rs.EOF And SelectedLocation = rs("Origin_Location_ID").Value Then
                        SelectedSectorNo = rs("SectorNo").Value
                        MsgBox "ya yao"
                        Me.ReleaseETD.Value = rs("Origin_ETD").Value
                        Me.DepartDate.Value = rs("Origin_DepartDate").Value
                        me.NextDestCmb.value=rs("DestAirportCode").Value
                    End If
                End If
            Elseif SelectedLocation = rs("Origin_Location_ID").Value Then
                SelectedSectorNo = rs("SectorNo").Value
                If SelectedSectorNo = 1 Then
                    Me.ReleaseETD.Value = rs("Origin_ETD").Value
                    Me.DepartDate.Value = rs("Origin_DepartDate").Value
                    Me.ReleaseETA.Value = "00:01"
                    Me.ArrivalDate.Value = rs("Origin_DepartDate").Value
                    me.NextDestCmb.value=rs("DestAirportCode").Value
                Else
                    Me.ReleaseETD.Value = rs("Origin_ETD").Value
                    Me.DepartDate.Value = rs("Origin_DepartDate").Value
                    me.NextDestCmb.value=rs("DestAirportCode").Value
                    SelectedSectorNo = rs("SectorNo").Value
                    If SelectedSectorNo > 1 Then
                        rs.MovePrevious
                        If Not rs.BOF Then
                            SelectedSectorNo = rs("SectorNo").Value
                            Me.ReleaseETA.Value = "Dest_ETA"
                            Me.ArrivalDate.Value = rs("Dest_ArrivalDate").Value
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
