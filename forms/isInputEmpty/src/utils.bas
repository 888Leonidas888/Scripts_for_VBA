Attribute VB_Name = "utils"
Public Function isInputsEmpty(frme As MSForms.frame) As Boolean

    Dim status As Boolean
    
    For Each ctrl In frme.Controls
        If TypeOf ctrl Is MSForms.TextBox And ctrl.Enabled Then
            If ctrl.Value = Empty Then
                ctrl.SetFocus
                status = True
                Exit For
            End If
        ElseIf TypeOf ctrl Is MSForms.ComboBox And ctrl.Enabled Then
            If ctrl.ListIndex = -1 Then
                ctrl.SetFocus
                status = True
                Exit For
            End If
        End If
    Next ctrl
    
    isInputsEmpty = status
    
End Function
Public Sub clearInputs(frme As MSForms.frame)

    For Each ctrl In frme.Controls
        If TypeOf ctrl Is MSForms.TextBox And ctrl.Enabled Then
            ctrl.Value = Empty
        ElseIf TypeOf ctrl Is MSForms.ComboBox And ctrl.Enabled Then
            ctrl.ListIndex = -1
        End If
    Next ctrl
    
End Sub
