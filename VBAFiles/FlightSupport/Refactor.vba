Private Sub ETD_LostFocus()
    Dim mainForm As Form
    Dim subForm1 As Form
    Dim subForm2 As Form
    Dim targetControl As control

    Forms!FlightSupport_LASTFORM_New_V12!FlightSupport_AddUpdateRequestF!Add_UpdateRequestF!FlightSupport_AddDestinationF

    'Set mainForm = Forms("FlightSupport_LASTFORM_New_V12")
    'Set subForm1 = mainForm.Controls("FlightSupport_AddUpdateRequestF").Form
    'Set subForm2 = subForm1.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_AddDestinationF").Form


    'Set mainForm = Forms("FlightSupport_LASTFORM_New_V12")
    'Set subForm1 = mainForm.Controls("FlightSupport_AddUpdateRequestF").Form
    'Set subForm2 = subForm1.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_AddDestinationF").Form

    ' Set focus To the subForm2 form


    ' Find the "Destination" control within the nested subforms
    Set targetControl = subForm2.Controls("Destination")

    If Not targetControl Is Nothing Then
        targetControl.SetFocus
    End If
End Sub


Private Sub ETD_LostFocus()
    Dim destForm As Form

    destForm=Forms!FlightSupport_LASTFORM_New_V12!FlightSupport_AddUpdateRequestF!Add_UpdateRequestF!FlightSupport_AddDestinationF
    destForm!Destination.SetFocus
End Sub