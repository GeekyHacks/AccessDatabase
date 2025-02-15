Private Sub CreateTableBtn_Click()
    On Error Goto ErrorHandler

        ' Initialize logging
        Call LogMessage("Process started.")

        ' Get the SQL code from the text box
        Dim SQLCode As String
        SQLCode = Me.SQLCode.Value
        Call LogMessage("SQL code retrieved from the text box. Value: " & SQLCode)

        ' Check If the text box is empty
        If Trim(SQLCode) = "" Then
            Call LogMessage("Error: SQL code is empty.")
            MsgBox "Please enter SQL code in the text box.", vbExclamation, "Error"
         Exit Sub
        End If

        ' Convert the SQL code To Access-compatible SQL
        Dim AccessSQL As String
        AccessSQL = ConvertToAccessSQL(SQLCode)
        Call LogMessage("SQL code converted To Access-compatible SQL. Value: " & AccessSQL)

        ' Extract the table name from the SQL statement
        Dim TableName As String
        TableName = ExtractTableName(AccessSQL)
        Call LogMessage("Table name extracted: " & TableName)

        ' Check If the table already exists
        Dim TableExists As Boolean
        TableExists = TableExistsInDatabase(TableName)
        If TableExists Then
            Call LogMessage("Table '" & TableName & "' already exists.")
        Else
            Call LogMessage("Table '" & TableName & "' does Not exist.")
        End If

        ' Show a warning If the table already exists
        If TableExists Then
            Dim Response As VbMsgBoxResult
            Response = MsgBox("The table '" & TableName & "' already exists. Do you want To overwrite it?", vbYesNo + vbExclamation, "Warning")
            If Response = vbNo Then
                Call LogMessage("User chose Not To overwrite the table.")
             Exit Sub
            Else
                ' Delete the existing table
                Call LogMessage("Deleting existing table '" & TableName & "'...")
                DoCmd.DeleteObject acTable, TableName
                Call LogMessage("Table '" & TableName & "' deleted successfully.")
                MsgBox "Table '" & TableName & "' deleted successfully!", vbInformation, "Success"
            End If
        End If

        ' Show a confirmation prompt before creating the table
        Dim Confirm As VbMsgBoxResult
        Confirm = MsgBox("Are you sure you want To create the table '" & TableName & "'?", vbYesNo + vbQuestion, "Confirmation")
        If Confirm = vbNo Then
            Call LogMessage("User chose Not To create the table.")
         Exit Sub
        End If

        ' Execute the SQL statement To create the table
        Call LogMessage("Creating table '" & TableName & "'...")

        ' Use CurrentDb.Execute instead of DoCmd.RunSQL
        CurrentDb.Execute AccessSQL
        Call LogMessage("Table '" & TableName & "' created successfully.")

        MsgBox "Table '" & TableName & "' created successfully!", vbInformation, "Success"

        ' Build relationships For the table
        Call LogMessage("Building relationships For table '" & TableName & "'...")
        BuildRelationships (AccessSQL)
        Call LogMessage("Relationships For table '" & TableName & "' built successfully.")

        ' Reset the form fields
        Call LogMessage("Resetting form fields...")
        ResetBtn_Click
        Call LogMessage("Form fields reset successfully.")

        Call LogMessage("Process completed successfully.")
     Exit Sub

 ErrorHandler:
        ' Log the error
        Call LogMessage("Error: " & Err.Description)
        ' Display an error message If something goes wrong
        MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub LogMessage(Message As String)
    ' Log messages To a text file
    Dim LogFilePath As String
    LogFilePath = "D:\Tahseen Work\TahseenPro\TahseenProAccess\VBAFiles\Debuglogs.txt"

    Dim FileNumber As Integer
    FileNumber = FreeFile

    ' Open the log file in append mode
    Open LogFilePath For Append As FileNumber
    ' Write the message With a timestamp
    Print #FileNumber, Now() & " - " & Message
    ' Close the log file
    Close FileNumber
End Sub

Private Function ConvertToAccessSQL(SQLCode As String) As String
    ' Convert non-Access SQL To Access-compatible SQL
    Dim AccessSQL As String
    AccessSQL = SQLCode

    ' Replace unsupported data types
    AccessSQL = Replace(AccessSQL, "VARCHAR", "TEXT")
    AccessSQL = Replace(AccessSQL, "NVARCHAR", "TEXT")
    AccessSQL = Replace(AccessSQL, "INT", "INTEGER")
    AccessSQL = Replace(AccessSQL, "BOOLEAN", "YESNO")
    AccessSQL = Replace(AccessSQL, "DECIMAL", "CURRENCY")
    AccessSQL = Replace(AccessSQL, "FLOAT", "DOUBLE")

    ' Replace AUTO_INCREMENT With COUNTER For Access compatibility
    AccessSQL = Replace(AccessSQL, "AUTOINCREMENT", "COUNTER")

    ' Remove the semicolon at the end of the SQL statement
    If Right(Trim(AccessSQL), 1) = ";" Then
        AccessSQL = Left(AccessSQL, Len(AccessSQL) - 1)
    End If

    ' Log the converted SQL
    Call LogMessage("Converted SQL: " & AccessSQL)

    ' Return the modified SQL code
    ConvertToAccessSQL = AccessSQL
End Function

Private Function TableExistsInDatabase(TableName As String) As Boolean
    ' Check If the table exists in the database
    On Error Resume Next
    TableExistsInDatabase = False
    Dim tdf As TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = TableName Then
            TableExistsInDatabase = True
         Exit For
        End If
    Next tdf
    On Error Goto 0
        ' Log the result
        Call LogMessage("TableExistsInDatabase: " & TableExistsInDatabase)
End Function

Private Function ExtractTableName(SQLStatement As String) As String
    ' Extract the table name from the SQL statement
    On Error Goto ExtractError
        Dim TableName As String
        Dim StartPos As Integer
        Dim EndPos As Integer

        ' Find the position of "CREATE TABLE"
        StartPos = InStr(1, SQLStatement, "CREATE TABLE", vbTextCompare)
        If StartPos = 0 Then
            ExtractTableName = ""
         Exit Function
        End If

        ' Find the position of the table name (after "CREATE TABLE")
        StartPos = StartPos + Len("CREATE TABLE")
        SQLStatement = Mid(SQLStatement, StartPos)

        ' Remove leading/trailing spaces And any brackets
        SQLStatement = Trim(SQLStatement)
        If Left(SQLStatement, 1) = "[" Then
            SQLStatement = Mid(SQLStatement, 2)
        End If
        If Right(SQLStatement, 1) = "]" Then
            SQLStatement = Left(SQLStatement, Len(SQLStatement) - 1)
        End If

        ' Extract the table name (up To the first space Or bracket)
        EndPos = InStr(1, SQLStatement, " ", vbTextCompare)
        If EndPos = 0 Then
            EndPos = InStr(1, SQLStatement, "(", vbTextCompare)
        End If
        If EndPos = 0 Then
            ExtractTableName = SQLStatement
        Else
            ExtractTableName = Left(SQLStatement, EndPos - 1)
        End If

        ' Log the extracted table name
        Call LogMessage("Extracted table name: " & ExtractTableName)

     Exit Function

 ExtractError:
        ExtractTableName = ""
        Call LogMessage("Error extracting table name: " & Err.Description)
End Function

Private Sub BuildRelationships(SQLStatement As String)
    On Error Goto ErrorHandler

        ' Extract foreign key constraints from the SQL statement
        Dim ForeignKeys As Collection
        Set ForeignKeys = ExtractForeignKeys(SQLStatement)

        ' Create relationships For each foreign key
        Dim i As Integer
        For i = 1 To ForeignKeys.Count
            Dim FK As Variant
            FK = ForeignKeys(i)

            ' Create the relationship
            CreateRelationship FK(0), FK(1), FK(2), FK(3)
        Next i

        ' Notify the user that the relationships were created successfully
        MsgBox "Relationships created successfully!", vbInformation, "Success"

     Exit Sub

 ErrorHandler:
        ' Display an error message If something goes wrong
        MsgBox "An error occurred While creating relationships: " & Err.Description, vbCritical, "Error"
        Call LogMessage("Error in BuildRelationships: " & Err.Description)
End Sub

Private Function ExtractForeignKeys(SQLStatement As String) As Collection
    ' Extract foreign key constraints from the SQL statement
    Dim ForeignKeys As Collection
    Set ForeignKeys = New Collection

    Dim Pos As Long
    Pos = InStr(1, SQLStatement, "FOREIGN KEY", vbTextCompare)
    Do While Pos > 0
        ' Extract the foreign key definition
        Dim FKDefinition As String
        Dim EndPos As Long
        EndPos = InStr(Pos, SQLStatement, ")", vbTextCompare)
        If EndPos > 0 Then
            FKDefinition = Mid(SQLStatement, Pos, EndPos - Pos + 1)
        End If

        ' Parse the foreign key definition
        Dim FKParts As Variant
        FKParts = Split(FKDefinition, "REFERENCES")
        If UBound(FKParts) = 1 Then
            Dim FKField As String
            FKField = Trim(Split(FKParts(0), "(")(1))
            FKField = Replace(FKField, ")", "")

            Dim RefTable As String
            RefTable = Trim(Split(FKParts(1), "(")(0))

            Dim RefField As String
            RefField = Trim(Split(FKParts(1), "(")(1))
            RefField = Replace(RefField, ")", "")

            ' Add the foreign key To the collection
            ForeignKeys.Add Array(FKField, RefTable, RefField, ExtractTableName(SQLStatement))
        End If

        ' Find the Next foreign key
        Pos = InStr(EndPos + 1, SQLStatement, "FOREIGN KEY", vbTextCompare)
    Loop

    ' Log the extracted foreign keys
    Call LogMessage("Extracted foreign keys: " & ForeignKeys.Count)

    ' Return the collection of foreign keys
    Set ExtractForeignKeys = ForeignKeys
End Function

Private Sub CreateRelationship(FKField As String, RefTable As String, RefField As String, TableName As String)
    On Error Goto ErrorHandler

        Dim db As DAO.Database
        Dim rel As DAO.Relation
        Dim fld As DAO.Field

        ' Get the current database
        Set db = CurrentDb

        ' Create the relationship
        Set rel = db.CreateRelation(TableName & "_" & RefTable, RefTable, TableName)
        rel.Attributes = dbRelationUpdateCascade + dbRelationDeleteCascade
        Set fld = rel.CreateField(FKField)
        fld.ForeignName = RefField
        rel.Fields.Append fld
        db.Relations.Append rel

        ' Log the created relationship
        Call LogMessage("Created relationship: " & TableName & "_" & RefTable)

     Exit Sub

 ErrorHandler:
        ' Display an error message If something goes wrong
        MsgBox "An error occurred While creating the relationship: " & Err.Description, vbCritical, "Error"
        Call LogMessage("Error in CreateRelationship: " & Err.Description)
End Sub

Private Sub ResetBtn_Click()
    ' Clear the text box
    Me.SQLCode.Value = Null
    Call LogMessage("ResetBtn_Click: Form fields reset.")
End Sub