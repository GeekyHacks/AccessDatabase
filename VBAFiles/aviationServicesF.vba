Option Compare Database

Private Sub closeBtn_Click()
DoCmd.Close acForm, "aviationServicesF"
End Sub

Private Sub Form_Load()
    DoCmd.GoToRecord , , acNewRec
End Sub

