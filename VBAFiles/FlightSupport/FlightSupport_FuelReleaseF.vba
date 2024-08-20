Private Sub secondarySectorID_AfterUpdate()
    Dim rsSector As DAO.Recordset
    Dim selectedSector As Integer
    Dim strSQL As String

    selectedSector = Me.secondarySectorID

    ' Build the SQL query With conditions
    strSQL = "Select Origin, Destination FROM sectortable WHERE ID=" & selectedSector

    ' Set the RowSource Property of the combobox
    Me.ReleaseLocation.RowSource = strSQL

    ' Requery the combobox
    Me.ReleaseLocation.Requery

End Sub


Private Sub ReleaseLocation_AfterUpdate()
    Dim rsSector As DAO.Recordset
    Dim selectedSector As Integer
    Dim selectedLocation As variant

    selectedSector = me.secondarySectorID
    selectedLocation= me.ReleaseLocation
    ' Open the sectortable recordset
    Set rsSector = CurrentDb.OpenRecordset("sectortable")

    ' Move To the first record in the sectortable
    rsSector.MoveFirst

    ' Loop through the sectortable records
    Do Until rsSector.EOF
        ' Check If the sector's Origin is EGGW
        If rsSector!ID = selectedSector
            ' Set the values For the fuel release record from the sectortable record
            If rsSector!Origin = selectedLocation Then
                me.DepartDate = rsSector!DepartDate
                me.ETD = rsSector!ETD
                me.Destination = rsSector!Destination
             Exit Do
            Else 
                ' Save the fuel release record
                me.ArrivalDate = rsSector!ArrivalDate
                me.ETA = rsSector!ETA
             Exit Do
            End If 
        End If

        ' Move To the Next record in the sectortable
        rsSector.MoveNext
    Loop

    ' Close the recordsets
    rsSector.Close
    Set rsSector = Nothing
End Sub


' using query For better preformance
Private Sub ReleaseLocation_AfterUpdate()
    Dim rsSector As DAO.Recordset
    Dim selectedSector As Integer
    Dim selectedLocation As Variant
    Dim strSQL As String

    selectedSector = Me.secondarySectorID
    selectedLocation = Me.ReleaseLocation

    ' Build the SQL query With conditions
    strSQL = "Select * FROM sectortable WHERE ID=" & selectedSector & " And (Origin='" & selectedLocation & "' Or Destination='" & selectedLocation & "')"

    ' Open the recordset With the SQL query
    Set rsSector = CurrentDb.OpenRecordset(strSQL)

    ' Check If any records are returned
    If Not rsSector.EOF Then
        ' Get the first record
        rsSector.MoveFirst

        ' Check If Origin matches selectedLocation
        If rsSector!Origin = selectedLocation Then
            ' Set the values For fuel release
            Me.DepartDate = rsSector!DepartDate
            Me.ETD = rsSector!ETD
            Me.Destination = rsSector!Destination
        Else
            ' Set the values For arrival date And ETA
            Me.ArrivalDate = rsSector!ArrivalDate
            Me.ETA = rsSector!ETA
        End If
    End If

    ' Close the recordset
    rsSector.Close
    Set rsSector = Nothing
End Sub