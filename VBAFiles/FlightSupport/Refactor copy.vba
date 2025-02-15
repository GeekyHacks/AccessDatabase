' create single table 

Private Sub CreateTableBtn_Click()
    On Error GoTo ErrorHandler

    ' Get the SQL code and table name from the text boxes
    Dim SQLCode As String
    Dim TableName As String
    SQLCode = Me.SQLCode.Value
    TableName = Me.TableName.Value

    ' Check if the text boxes are empty
    If Trim(SQLCode) = "" Or Trim(TableName) = "" Then
        MsgBox "Please enter both SQL code and table name.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Check if the table already exists
    Dim TableExists As Boolean
    TableExists = False
    Dim tdf As TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = TableName Then
            TableExists = True
            Exit For
        End If
    Next tdf

    ' Show a warning if the table already exists
    If TableExists Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("The table '" & TableName & "' already exists. Do you want to overwrite it?", vbYesNo + vbExclamation, "Warning")
        If Response = vbNo Then
            Exit Sub
        End If
    End If

    ' Show a confirmation prompt before creating the table
    Dim Confirm As VbMsgBoxResult
    Confirm = MsgBox("Are you sure you want to create the table '" & TableName & "'?", vbYesNo + vbQuestion, "Confirmation")
    If Confirm = vbNo Then
        Exit Sub
    End If

    ' Execute the SQL code to create the table
    DoCmd.RunSQL SQLCode

    ' Notify the user that the table was created successfully
    MsgBox "Table created successfully!", vbInformation, "Success"

    ' Reset the form fields
    ResetBtn_Click

    Exit Sub

ErrorHandler:
    ' Display an error message if something goes wrong
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ResetBtn_Click()
    ' Clear the text boxes
    Me.SQLCode.Value = Null
    Me.TableName.Value = Null
End Sub