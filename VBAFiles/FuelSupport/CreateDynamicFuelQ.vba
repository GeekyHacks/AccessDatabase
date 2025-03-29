Option Compare Database
Option Explicit
'=======================================================
' Enhanced Logging System
'=======================================================
Private Sub LogError( _
    ByVal ErrorMessage As String, _
    ByVal ErrorSource As String, _
    Optional ByVal ErrorNumber As Long = 0, _
    Optional ByVal ModuleName As String = "", _
    Optional ByVal ProcedureName As String = "", _
    Optional ByVal AdditionalInfo As String = "" _
    )
    On Error GoTo LogError_Handler
        Dim filePath As String
        Dim fileNumber As Integer
        Dim logMessage As String

        logMessage = String(50, "=") & vbCrLf & _
        "Timestamp: " & Now() & vbCrLf & _
        "Module: " & ModuleName & vbCrLf & _
        "Procedure: " & ProcedureName & vbCrLf & _
        "Error #: " & ErrorNumber & vbCrLf & _
        "Source: " & ErrorSource & vbCrLf & _
        "Message: " & ErrorMessage & vbCrLf & _
        "Additional Info: " & AdditionalInfo & vbCrLf & _
        String(50, "=") & vbCrLf & vbCrLf

        filePath = CurrentProject.Path & "\addfuelquote.log"
        fileNumber = FreeFile
        Open filePath For Append As #fileNumber
        Print #fileNumber, logMessage
        Close #fileNumber
     Exit Sub

LogError_Handler:
        MsgBox "Logging System Failure: " & Err.Description, vbCritical
End Sub

'=======================================================
' Main Fuel Table Creation Function
'=======================================================
Public Function CreateDynamicFuelQuery() As Boolean
    On Error GoTo ErrorHandler
        Const PROC_NAME As String = "CreateDynamicFuelQuery"

        Dim db As DAO.Database
        Dim rsSource As DAO.Recordset
        Dim rsTemp As DAO.Recordset
        Dim qdf As DAO.QueryDef
        Dim strSQL As String
        Dim tblName As String
        Dim successCount As Integer
        Dim tempRecordCount As Long
        Dim tempVendorCount As Long

        ' Initialize logging
        LogError "Function started", "Info", 0, "FuelModule", PROC_NAME

        Set db = CurrentDb()
        LogError "Database connection established", "Info", 0, "FuelModule", PROC_NAME

        ' Get all vendor tables from aviationServicesT_EX
        LogError "Retrieving vendor tables from aviationServicesT_EX", "Info", 0, "FuelModule", PROC_NAME
        Set rsSource = db.OpenRecordset("Select FuelTablesList FROM aviationServicesT_EX")

        If rsSource.EOF Then
            LogError "No vendor tables found in aviationServicesT_EX", "Data Error", 0, "FuelModule", PROC_NAME
            MsgBox "No fuel tables found in aviationServicesT_EX!", vbExclamation
            CreateDynamicFuelQuery = False
            GoTo Cleanup
            End If

            ' Create temporary table
            On Error Resume Next
            db.Execute "DROP TABLE FuelSupport_FuelPricesT_V12", dbFailOnError
            If Err.Number <> 0 Then
                LogError "Error dropping temp table: " & Err.Description, "Database", Err.Number, "FuelModule", PROC_NAME
                Err.Clear
            End If

            On Error GoTo ErrorHandler
                db.Execute "CREATE TABLE FuelSupport_FuelPricesT_V12 (" & _
                "Airport TEXT(55), " & _
                "FlightType TEXT(55), " & _
                "VendorPrice DOUBLE, " & _
                "TableName TEXT(100), " & _
                "ID LONG, " & _
                "LOWCURRENCY TEXT(50), " & _
                "EstimatedTotal DOUBLE, " & _
                "Validity DATE, " & _
                "EffectivePrice DOUBLE)"
                LogError "Temporary table created successfully", "Info", 0, "FuelModule", PROC_NAME

                ' Process each vendor table
                successCount = 0
                Do Until rsSource.EOF
                    tblName = Trim(Nz(rsSource!FuelTablesList, ""))

                    If tblName = "" Then
                        LogError "Empty table name encountered", "Warning", 0, "FuelModule", PROC_NAME
                        GoTo NextRecord
                        End If

                        If Not TableExists(db, tblName) Then
                            LogError "Table '" & tblName & "' does Not exist", "Validation", 0, "FuelModule", PROC_NAME
                            GoTo NextRecord
                            End If

                            ' Insert data into temp table
                            On Error Resume Next
                            strSQL = "INSERT INTO FuelSupport_FuelPricesT_V12 " & _
                            "(Airport, FlightType, VendorPrice, TableName, ID, " & _
                            "LOWCURRENCY, EstimatedTotal, Validity, EffectivePrice) " & _
                            "Select Airport, FlightType, VendorPrice, '" & tblName & "', " & _
                            "ID, Currency, EstimatedTotal, Validity, " & _
                            "IIf(VendorPrice > EstimatedTotal, VendorPrice, EstimatedTotal) " & _
                            "FROM [" & tblName & "]"

                            db.Execute strSQL, dbFailOnError

                            If Err.Number = 0 Then
                                successCount = successCount + 1
                                LogError "Successfully inserted data from: " & tblName, "Info", 0, "FuelModule", PROC_NAME
                            Else
                                LogError "Failed To insert from " & tblName & ": " & Err.Description, "Error", Err.Number, "FuelModule", PROC_NAME
                                Err.Clear
                            End If

NextRecord:
                            rsSource.MoveNext
                        Loop

                        ' Validate data insertion
                        If successCount = 0 Then
                            LogError "No data inserted from any vendor tables", "Critical", 0, "FuelModule", PROC_NAME
                            MsgBox "Failed To load any vendor data!", vbCritical
                            CreateDynamicFuelQuery = False
                            GoTo Cleanup
                            End If

                            ' Verify temp table contents
                            Set rsTemp = db.OpenRecordset("Select COUNT(*) As RecCount FROM FuelSupport_FuelPricesT_V12")
                            tempRecordCount = rsTemp!RecCount
                            rsTemp.Close

                            Set rsTemp = db.OpenRecordset("Select COUNT(DISTINCT TableName) As VendorCount FROM FuelSupport_FuelPricesT_V12")
                            tempVendorCount = rsTemp!vendorCount
                            rsTemp.Close

                            LogError "Temp table contains " & tempRecordCount & " records from " & tempVendorCount & " vendors", _
                            "Info", 0, "FuelModule", PROC_NAME

                            CreateDynamicFuelQuery = True
                            LogError "Function completed successfully", "Info", 0, "FuelModule", PROC_NAME

Cleanup:
                            On Error Resume Next
                            rsSource.Close
                            Set rsSource = Nothing
                            Set rsTemp = Nothing
                            Set qdf = Nothing
                            Set db = Nothing
                         Exit Function

ErrorHandler:
                            LogError Err.Description, "Runtime Error", Err.Number, "FuelModule", PROC_NAME, "Error occurred at step: " & Erl()
                            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
                            CreateDynamicFuelQuery = False
                            Resume Cleanup
End Function
'=======================================================
' Valid Fuel Prices Table Creation
'=======================================================
Public Function CreateValidFuelPricesTable() As Boolean
    On Error GoTo ErrorHandler
        Const PROC_NAME As String = "CreateValidFuelPricesTable"

        Dim db As DAO.Database
        Dim rsSource As DAO.Recordset
        Dim rsTemp As DAO.Recordset
        Dim strSQL As String
        Dim tblName As String
        Dim successCount As Integer
        Dim tempRecordCount As Long

        ' Initialize logging
        LogError "Function started", "Info", 0, "FuelModule", PROC_NAME

        Set db = CurrentDb()
        LogError "Database connection established", "Info", 0, "FuelModule", PROC_NAME

        ' Get all vendor tables from aviationServicesT_EX
        LogError "Retrieving vendor tables from aviationServicesT_EX", "Info", 0, "FuelModule", PROC_NAME
        Set rsSource = db.OpenRecordset("Select FuelTablesList FROM aviationServicesT_EX")

        If rsSource.EOF Then
            LogError "No vendor tables found", "Data Error", 0, "FuelModule", PROC_NAME
            MsgBox "No fuel tables found in aviationServicesT_EX!", vbExclamation
            CreateValidFuelPricesTable = False
            GoTo Cleanup
            End If

            ' Create validated prices table
            On Error Resume Next
            db.Execute "DROP TABLE FuelSupport_ValidFuelPricesT_V12", dbFailOnError
            If Err.Number <> 0 Then
                LogError "Error dropping old table: " & Err.Description, "Database", Err.Number, "FuelModule", PROC_NAME
                Err.Clear
            End If

            On Error GoTo ErrorHandler
                db.Execute "CREATE TABLE FuelSupport_ValidFuelPricesT_V12 (" & _
                "Airport TEXT(55), " & _
                "FlightType TEXT(55), " & _
                "VendorPrice DOUBLE, " & _
                "TableName TEXT(100), " & _
                "ID LONG, " & _
                "LOWCURRENCY TEXT(50), " & _
                "Unit TEXT(10), " & _
                "EstimatedTotal DOUBLE, " & _
                "Validity DATE, " & _
                "MinPrice DOUBLE)"
                LogError "Valid prices table created", "Info", 0, "FuelModule", PROC_NAME

                ' Process each vendor table With validity check
                successCount = 0
                Do Until rsSource.EOF
                    tblName = Trim(Nz(rsSource!FuelTablesList, ""))

                    If tblName = "" Then
                        LogError "Empty table name skipped", "Warning", 0, "FuelModule", PROC_NAME
                        GoTo NextRecord
                        End If

                        If Not TableExists(db, tblName) Then
                            LogError "Missing table: " & tblName, "Validation", 0, "FuelModule", PROC_NAME
                            GoTo NextRecord
                            End If

                            ' Insert only valid records (validity >= today)
                            On Error Resume Next
                            strSQL = "INSERT INTO FuelSupport_ValidFuelPricesT_V12 " & _
                            "(Airport, FlightType, VendorPrice, TableName, ID, " & _
                            "LOWCURRENCY, Unit, EstimatedTotal, Validity, MinPrice) " & _
                            "Select Airport, FlightType, VendorPrice, '" & tblName & "', " & _
                            "ID, Currency, Unit, EstimatedTotal, Validity, " & _
                            "IIf(VendorPrice > EstimatedTotal, VendorPrice, EstimatedTotal) " & _
                            "FROM [" & tblName & "] " & _
                            "WHERE Validity >= Date()" ' Validity filter here

                            db.Execute strSQL, dbFailOnError

                            If Err.Number = 0 Then
                                successCount = successCount + 1
                                LogError "Added valid records from: " & tblName, "Success", 0, "FuelModule", PROC_NAME
                            Else
                                LogError "Failed To add from " & tblName & ": " & Err.Description, "Error", Err.Number, "FuelModule", PROC_NAME
                                Err.Clear
                            End If

NextRecord:
                            rsSource.MoveNext
                        Loop

                        ' Validate results
                        If successCount = 0 Then
                            LogError "No valid data inserted", "Critical", 0, "FuelModule", PROC_NAME
                            MsgBox "No valid fuel prices found!", vbCritical
                            CreateValidFuelPricesTable = False
                            GoTo Cleanup
                            End If

                            ' Verify table contents
                            Set rsTemp = db.OpenRecordset("Select COUNT(*) As ValidCount FROM FuelSupport_ValidFuelPricesT_V12")
                            tempRecordCount = rsTemp!ValidCount
                            rsTemp.Close

                            LogError "Valid table contains " & tempRecordCount & " records", "Info", 0, "FuelModule", PROC_NAME
                            MsgBox "Valid fuel prices table created With " & tempRecordCount & " current entries!", vbInformation

                            CreateValidFuelPricesTable = True
                            LogError "Function completed successfully", "Info", 0, "FuelModule", PROC_NAME

Cleanup:
                            On Error Resume Next
                            rsSource.Close
                            Set rsSource = Nothing
                            Set rsTemp = Nothing
                            Set db = Nothing
                         Exit Function

ErrorHandler:
                            LogError Err.Description, "Runtime Error", Err.Number, "FuelModule", PROC_NAME, "Error at step: " & Erl()
                            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
                            CreateValidFuelPricesTable = False
                            Resume Cleanup
End Function
'=======================================================
' Helper Functions
'=======================================================
'=======================================================
' Helper Functions
'=======================================================
Public Function TableExists(db As DAO.Database, TableName As String) As Boolean
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "TableExists"
    
    Dim tdf As DAO.TableDef
    
    ' Log function start
    LogError "Checking existence of table: " & TableName, "Info", 0, "FuelModule", PROC_NAME
    
    Set tdf = db.TableDefs(TableName)
    TableExists = True
    
    ' Log success
    LogError "Table '" & TableName & "' exists", "Info", 0, "FuelModule", PROC_NAME
    Exit Function
    
ErrorHandler:
    TableExists = False
    ' Log the error/non-existence
    LogError "Table '" & TableName & "' does not exist", "Info", Err.Number, "FuelModule", PROC_NAME
    Resume Next
End Function