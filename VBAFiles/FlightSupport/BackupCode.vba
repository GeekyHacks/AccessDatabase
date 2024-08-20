Private Sub BackUpBtn_Click()
    ExportAllTablesToExcel
End Sub

Public Sub ExportAllTablesToExcel()
    Dim db As DAO.Database
    Dim tblDef As DAO.TableDef
    Dim exportPath As String
    Dim excelFileName As String

    ' Set the directory where you want To save the Excel files
    exportPath = Environ("USERPROFILE") & "\Documents\TahseenPro_BackUp\" ' Change this To your actual directory path

    ' Check If the directory exists, If Not, create it
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    ' Set the database variable To the current database
    Set db = CurrentDb()

    ' Loop through all table definitions in the database
    For Each tblDef In db.TableDefs
        ' Ignore system And temporary tables
        If Not (tblDef.Name Like "MSys*" Or tblDef.Name Like "~*") Then
            ' Define the file name For Excel format
            excelFileName = exportPath & tblDef.Name & ".xlsx"

            ' Export the table To Excel
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tblDef.Name, excelFileName, True
        End If
    Next tblDef

    ' Clean up
    Set tblDef = Nothing
    Set db = Nothing
    ExportQuerySQLToExcel

    MsgBox "tables & SQL code exported To Excel files in " & exportPath, vbInformation
End Sub


Private Sub ExportQuerySQLToExcel()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim exportPath As String
    Dim queryName As String
    Dim sqlCode As String
    ' Set the directory where you want To save the Excel files
    exportPath = Environ("USERPROFILE") & "\Documents\TahseenPro_BackUp\Queries\"

    ' Check If the directory exists, If Not, create it
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    ' Set the database variable To the current database
    Set db = CurrentDb()

    ' Create an Excel application
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False

    ' Loop through all query definitions in the database
    For Each qdf In db.QueryDefs
        queryName = qdf.Name
        sqlCode = qdf.SQL

        ' Create a New Excel workbook
        Set excelWorkbook = excelApp.Workbooks.Add
        Set excelWorksheet = excelWorkbook.Sheets(1)

        ' Write query name And SQL code To the worksheet
        excelWorksheet.Cells(1, 1).Value = "Query Name"
        excelWorksheet.Cells(1, 2).Value = "SQL Code"
        excelWorksheet.Cells(2, 1).Value = queryName
        excelWorksheet.Cells(2, 2).Value = sqlCode

        ' Save the workbook
        excelWorkbook.SaveAs exportPath & queryName & ".xlsx"
        excelWorkbook.Close SaveChanges:=False
    Next qdf

    ' Clean up
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    excelApp.Quit
    Set excelApp = Nothing
    Set qdf = Nothing
    Set db = Nothing


End Sub