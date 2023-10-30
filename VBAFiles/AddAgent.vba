Option Compare Database

Private Sub closeBtn_Click()
DoCmd.Close acForm, "AddAgentF"
End Sub

Private Sub Form_Load()
    DoCmd.GoToRecord , , acNewRec
End Sub

