Option Compare Database
Private Function GetSubreportRecordCount(subreportName As String) As Long
    Dim subreport As Report
    Dim recordCount As Long

    On Error Resume Next
    Set subreport = Me.Controls(subreportName).Report

    If Not subreport Is Nothing Then
        recordCount = DCount("*", subreport.RecordSource)
    End If

    GetSubreportRecordCount = recordCount
End Function
Private Sub FlightSupport_Schedule_New_Enter()

    ' Loop through each field in the recordset
    ' Check If the field name matches the specified fields
    If Me.DepartDate = date Or Me.DepartDate > date Then
        Me.SectorNo.ForeColor = RGB(151, 148, 194)
        Me.DepartDate.ForeColor = RGB(151, 148, 194)
        Me.ETD.ForeColor = RGB(151, 148, 194)
        Me.ETA.ForeColor = RGB(151, 148, 194)
        Me.ArrivalDate.ForeColor = RGB(151, 148, 194)
        Me.ORIGIN.ForeColor = RGB(151, 148, 194)
        Me.Destination.ForeColor = RGB(151, 148, 194)
        Me.CallSign.ForeColor = RGB(151, 148, 194)
        Me.SectorNo.FontItalic = True
        Me.DepartDate.FontItalic = True
        Me.ETD.FontItalic = True
        Me.ETA.FontItalic = True
        Me.ArrivalDate.FontItalic = True
        Me.ORIGIN.FontItalic = True
        Me.Destination.FontItalic = True
        Me.CallSign.FontItalic = True
    Else
        Me.SectorNo.FontBold = True
        Me.DepartDate.FontBold = True
        Me.ETD.FontBold = True
        Me.ETA.FontBold = True
        Me.ArrivalDate.FontBold = True
        Me.ORIGIN.FontBold = True
        Me.Destination.FontBold = True
        Me.CallSign.FontBold = True
    End If
End Sub
Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    Cancel = Not Me.ConciergeBrief_R.Report.HasData
End Sub
Private Sub GroupFooter6_Format(Cancel As Integer, FormatCount As Integer)
    Cancel = Not Me.FuelBrief_R.Report.HasData
End Sub
Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)
    ' Get the value of the Location field For the current record in the group
    Dim locationValue As Variant
    locationValue = DLookup("Location", "[FlightSupport_ExternalReport_New]", "Nz([Sect_ID], '') = '" & CStr(Me.[Sect_ID]) & "'")

    ' Get the value of the Tripstatus field For the current record in the group
    Dim statusValue As Variant
    statusValue = DLookup("Tripstatus", "[FlightSupport_ExternalReport_New]", "Nz([Sect_ID], '') = '" & CStr(Me.[Sect_ID]) & "'")

    ' Check If either condition is True
    If IsNull(locationValue) Or statusValue = "CANCELED" Then
        Cancel = True ' Stop the group header from being rendered
    Else
        ' If both conditions are False, make the field visible
        Cancel = False
    End If

End Sub
Private Sub GroupHeader3_Format(Cancel As Integer, FormatCount As Integer)
    Cancel = Not Me.PermitBrief_R.Report.HasData

End Sub
Private Sub GroupHeader4_Format(Cancel As Integer, FormatCount As Integer)
    If IsNull(Me.Locationtxt) Then
        Cancel = IsNull(Me.Locationtxt)
    End If
End Sub
Private Sub GroupHeader5_Format(Cancel As Integer, FormatCount As Integer)

    Cancel = Not Me.HandlingBrief_R.Report.HasData

End Sub
Private Sub Report_Open(Cancel As Integer)
    ' Set the RecordSource of the report To "FlightSupportQ"
    Me.RecordSource = "FlightSupport_ExternalReport_New"
End Sub

