Option Compare Database

Private Sub AddACbtn_Click()
    DoCmd.OpenForm "ACRegF", WindowMode:=acWindowNormal
End Sub

Private Sub AddAirportBtn_Click()
    DoCmd.OpenForm "AddAirports_FIRsF", WindowMode:=acWindowNormal
End Sub

Private Sub addAServiceBtn_Click()
    DoCmd.OpenForm "aviationServicesF", WindowMode:=acWindowNormal
End Sub

Private Sub AddICAO_IATABtn_Click()
    DoCmd.OpenForm "AddNewAirportF", WindowMode:=acWindowNormal
End Sub

Private Sub closeBtn_Click()
    Me.Undo
    DoCmd.Close acForm, "LASTFORM", acSaveNo
End Sub

Private Sub Command34_Click()
    DoCmd.OpenForm "FuelSupport_FuelQuote", WindowMode:=acWindowNormal
End Sub

Private Sub Command35_Click()
    DoCmd.OpenForm "FuelSupport_AddFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command36_Click()
    DoCmd.OpenForm "FuelSupport_UpdateFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command41_Click()
    DoCmd.OpenForm "FlightSupport_ReportGeneratorF", WindowMode:=acDialog
End Sub

Private Sub Command42_Click()
    DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acDialog
End Sub

Private Sub Command49_Click()
    DoCmd.OpenForm "FlightSupport_UrgentRequests", WindowMode:=acDialog
End Sub

Private Sub Detail_Click()
    Me.Undo
End Sub

Private Sub Form_Load()
    Dim fw As New clFormWindow

    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
End Sub

Private Sub IATA_CODE_Change()
    Dim selectedValue As Variant
    selectedValue = Me.IATA_CODE.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("Airports_DataT", dbOpenSnapshot)

        rs.FindFirst "IATA = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ICAO_CODE.Value = rs("ICAO").Value
            Me.CountryCmb.Value = rs("Country").Value
        Else
            Me.ICAO_CODE.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ICAO_CODE.Value = Null
    End If

 Exit Sub

End Sub
Private Sub RefreshBtn_Click()
    Me.Requery
    Me.Refresh
End Sub

Private Sub UpdateFormBtn_Click()
    DoCmd.OpenForm "FlightSupport_UpdateLASTFORM", WindowMode:=acWindowNormal
End Sub

Private Sub UpdateFuelReleasebtn_Click()
    DoCmd.OpenForm "FlightSupport_UpdateFuelReleaseF", WindowMode:=acWindowNormal
End Sub
///////////////////////////////////////////////////////////////////////////////

'OLD CODE 

Option Compare Database

Private Sub AddACbtn_Click()
    DoCmd.OpenForm "ACRegF", WindowMode:=acWindowNormal
End Sub

Private Sub AddAirportBtn_Click()
    DoCmd.OpenForm "AddAirports_FIRsF", WindowMode:=acWindowNormal
End Sub

Private Sub addAServiceBtn_Click()
    DoCmd.OpenForm "aviationServicesF", WindowMode:=acWindowNormal
End Sub

Private Sub AddICAO_IATABtn_Click()
    DoCmd.OpenForm "AddNewAirportF", WindowMode:=acWindowNormal
End Sub

Private Sub closeBtn_Click()

    Me.Undo

    DoCmd.Close acForm, "LASTFORM", acSaveNo

End Sub
Private Sub Command34_Click()
    DoCmd.OpenForm "FuelSupport_FuelQuote", WindowMode:=acWindowNormal
End Sub

Private Sub Command35_Click()
    DoCmd.OpenForm "FuelSupport_AddFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command36_Click()
    DoCmd.OpenForm "FuelSupport_UpdateFuel", WindowMode:=acWindowNormal
End Sub

Private Sub Command41_Click()
    DoCmd.OpenForm "FlightSupport_ReportGeneratorF", WindowMode:=acDialog
End Sub

Private Sub Command42_Click()
    DoCmd.OpenForm "FuelSupport_ReportGeneratorF", WindowMode:=acDialog
End Sub

Private Sub Command49_Click()
    DoCmd.OpenForm "FlightSupport_UrgentRequests", WindowMode:=acDialog
End Sub

Private Sub Detail_Click()
    Me.Undo
End Sub

Private Sub Form_Load()
    Dim fw As New clFormWindow

    fw.hWnd = Me.hWnd
    With fw
        .Top = (.Parent.Height - .Height) / 2
        .Left = (.Parent.Width - .Width) / 2
    End With
    Set fw = Nothing
End Sub

Private Sub IATA_CODE_Change()
    Dim selectedValue As Variant
    selectedValue = Me.IATA_CODE.Value ' Get the selected value from the combo box

    If Not IsNull(selectedValue) Then
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("Airports_DataT", dbOpenSnapshot)

        rs.FindFirst "IATA = '" & selectedValue & "'"
        If Not rs.NoMatch Then
            Me.ICAO_CODE.Value = rs("ICAO").Value
            Me.CountryCmb.Value = rs("Country").Value
        Else
            Me.ICAO_CODE.Value = Null
        End If

        rs.Close
        Set rs = Nothing
    Else
        Me.ICAO_CODE.Value = Null
    End If

 Exit Sub

End Sub
Private Sub RefreshBtn_Click()
    Me.Requery
    Me.Refresh
End Sub

Private Sub UpdateFormBtn_Click()
    DoCmd.OpenForm "FlightSupport_UpdateLASTFORM", WindowMode:=acWindowNormal
End Sub

Private Sub UpdateFuelReleasebtn_Click()
    DoCmd.OpenForm "FlightSupport_UpdateFuelReleaseF", WindowMode:=acWindowNormal
End Sub
