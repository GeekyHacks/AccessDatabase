Private Sub btnRefresh_Click()
    On Error Goto ErrorHandler

        ' Set the current form
        Dim CurrentForm As Form
        Set CurrentForm = Forms("FlightSupport_LASTFORM_New_V12").Form

        ' Requery the main form's subforms
        CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Requery
        CurrentForm.Controls("FlightSupport_RequestF_New").Form.Requery

        ' Requery all combo boxes in the main form
        RequeryComboBoxes CurrentForm

        ' Requery all combo boxes in the subform FlightSupport_RequestF_New
        RequeryComboBoxes CurrentForm.Controls("FlightSupport_RequestF_New").Form

        ' Requery all combo boxes in the subform FlightSupport_AddUpdateRequestF
        RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form

        ' Requery all combo boxes in the nested subforms under FlightSupport_AddUpdateRequestF
        RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form

        ' Requery all combo boxes in the subforms under navigationbutton12 If they are visible
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_AddOriginF_Form") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_AddOriginF_Form").Form
        End If
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_AddDestinationF") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_AddDestinationF").Form
        End If

        ' Requery all combo boxes in the subforms under navigationbutton20 If they are visible
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_UpdateSectorsF_New") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_UpdateSectorsF_New").Form
        End If
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_UpdateFuelBriefF") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_UpdateFuelBriefF").Form
        End If
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_UpdateFuelReleaseF_New") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_UpdateFuelReleaseF_New").Form
        End If
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_PermitsListF_New") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_PermitsListF_New").Form
        End If
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_HandlingUpdateF") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_HandlingUpdateF").Form
        End If
        If IsControlVisible(CurrentForm, "FlightSupport_AddUpdateRequestF", "Add_UpdateRequestF", "FlightSupport_ConciergeBriefUpdateF") Then
            RequeryComboBoxes CurrentForm.Controls("FlightSupport_AddUpdateRequestF").Form.Controls("Add_UpdateRequestF").Form.Controls("FlightSupport_ConciergeBriefUpdateF").Form
        End If

        ' Refresh the main form
        CurrentForm.Refresh

     Exit Sub

 ErrorHandler:
        ' Handle the error And continue
     Resume Next
End Sub

Private Sub RequeryComboBoxes(frm As Form)
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If ctrl.ControlType = acComboBox Then
            ctrl.Requery
        End If
        ' Check If the control is a subform And requery its combo boxes
        If ctrl.ControlType = acSubform Then
            RequeryComboBoxes ctrl.Form
        End If
    Next ctrl
End Sub

Private Function IsControlVisible(frm As Form, ParamArray ctrlNames() As Variant) As Boolean
    On Error Goto ErrorHandler
        Dim ctrl As Control
        Set ctrl = frm
        Dim i As Integer
        For i = LBound(ctrlNames) To UBound(ctrlNames)
            Set ctrl = ctrl.Controls(ctrlNames(i))
        Next i
        IsControlVisible = ctrl.Visible
     Exit Function

 ErrorHandler:
        IsControlVisible = False
End Function