Private Sub cmdGenerate_Click()
    On Error GoTo ErrorHandler

    Dim rst As DAO.Recordset
    Dim randomRefValue As Integer
    Dim RefNo As String
    Dim isUnique As Boolean

    Do
        ' Generate a random number between 1 and 9999
        randomRefValue = Int((9999 - 1 + 1) * Rnd + 1)

        ' Check if the generated number is greater than or equal to 176
        If randomRefValue >= 176 Then
            ' Generate the reference number
            RefNo = "TAS-" & Format(randomRefValue, "0000") & "_" & Format(Date, "yy")

            ' Check if the generated reference number already exists in the FlightSupport_FlightRequestsT table
            Set rst = CurrentDb.OpenRecordset("SELECT RefNo FROM FlightSupport_FlightRequestsT WHERE RefNo = '" & RefNo & "'", dbOpenSnapshot)

            If rst.EOF Then
                ' If the generated reference number does not exist, assign it to the form field
                Me.RefNo.Value = RefNo

                ' Set the flag to indicate that a unique reference number has been generated
                isUnique = True
            End If

            rst.Close
            Set rst = Nothing
        End If
    Loop Until isUnique

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    If Not rst Is Nothing Then
        rst.Close
        Set rst = Nothing
    End If
End Sub 