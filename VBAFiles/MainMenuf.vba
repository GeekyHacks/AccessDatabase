Option Compare Database

Private Sub addAServiceBtn_Click()
DoCmd.OpenForm "aviationServicesF", WindowMode:=acDialog
End Sub

Private Sub FlightSupportBtn_Click()
    DoCmd.OpenForm "FlightRequestF", WindowMode:=acDialog
    
End Sub

Private Sub closeBtn_Click()
DoCmd.Close acForm, "MainMenuF"
End Sub

Private Sub Command34_Click()
DoCmd.OpenForm "FlightRequestF", WindowMode:=acDialog
End Sub
