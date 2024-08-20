
Option Compare Database

'Private Sub Form_Load()
    'LoadLastBtn_Click
    'End Sub

Private Sub LoadLastBtn_Click()
    Me.RefNo.DefaultValue = "'" & GetLastRefNo() & "'"
End Sub

Private Sub newSectorBtn_Click()
    ' Clear the form fields To prepare For a New record
    cmdGenerate_Click
    ' DoCmd.RunCommand acCmdRecordsGoToNew
    'Me.RefNo.Value = GetLastRefNo
    ' Me.SectorNo.Value = ""
    '  Me.SectorRoute.Value = ""
    ' Me.Location.Value = ""
    ' Me.SectorNo.SetFocus
End Sub

Private Sub cmdGenerate_Click()
    Dim RefValue As Integer
    Dim rst As DAO.Recordset

    RefValue = Nz(DLookup("RefNum", "TSect_ID"), 0)

    Me.Sect_ID = Format(RefValue + 1, "000")

    Set rst = CurrentDb.OpenRecordset("TSect_ID", dbOpenDynaset, dbSeeChanges)

    With rst
        .Edit
        ![RefNum] = ![RefNum] + 1
        .Update
    End With

    rst.Close
    Set rst = Nothing

End Sub
Private Function GetLastRefNo() As Variant

    ' Get the latest RefNo value from NewFlightRequestsT table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    'Dim lastRefNo As Variant
    Dim MaxID As Long
    MaxID = DMax("ID", "NewSectorsT")
    ' Me.RecordSource = "Select * FROM NewSectorsT WHERE ID = " & MaxID & ";"

    Set db = CurrentDb
    strSQL = "Select * FROM NewSectorsT WHERE ID = " & MaxID & ";"
    Set rs = db.OpenRecordset(strSQL)

    ' Check If a record was found
    If Not rs.EOF Then
        lastRefNo = rs.Fields("RefNo").Value
    End If

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLastRefNo = lastRefNo
End Function

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Me.SectorNo.Value = "" Or Me.SectorRoute.Value = "" Or Me.Location.Value = "" Then
        MsgBox "Please fill in all required fields.", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub AddSectorBtn_Click()
    If IsNull(Me.RefNo.Value) Or IsNull(Me.Sect_ID.Value) Or IsNull(Me.SectorNo.Value) Or IsNull(Me.SectorRoute.Value) Or IsNull(Me.Location.Value) Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Incomplete Data"
    Else
        DoCmd.RunCommand acCmdSaveRecord
        DoCmd.RunCommand acCmdRecordsGoToNew
        cmdGenerate_Click
        Me.RefNo.Value = GetLastRefNo
    End If

End Sub

Public Sub ResetBtn_Click()
    On Error Goto ErrorHandler

        Me.Undo
        DoCmd.CancelEvent
        DoCmd.GoToRecord , , acNewRec

     Exit Sub

 ErrorHandler:
        MsgBox "An error occurred: " & err.Description, vbExclamation
End Sub


