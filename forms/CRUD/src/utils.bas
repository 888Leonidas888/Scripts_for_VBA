Attribute VB_Name = "utils"
Sub startApp()

    frmShowBooks.Show
    
End Sub

Public Sub clearInputs(frme As MSForms.Frame)

    For Each ctrl In frme.Controls
        If TypeOf ctrl Is MSForms.TextBox And ctrl.Enabled Then
            ctrl.value = Empty
        ElseIf TypeOf ctrl Is MSForms.ComboBox And ctrl.Enabled Then
            ctrl.ListIndex = -1
        End If
    Next ctrl
    
End Sub
Public Sub settingFrmAddBooks(frm As UserForm)
    
    With frm
        .Caption = "Record book"
        .Frame1.Caption = "New book"
    End With
    
    
    With frm.cmbBadge
        .AddItem "PEN"
        .AddItem "COP"
        .AddItem "DOL"
        .AddItem "LIB"
    End With
    
    With frm
        .txtCreated.value = Date
        .txtUpdate.value = Date
    End With
    
    With frm
        .btnAddBook.ControlTipText = "Add"
        .btnClear.ControlTipText = "Clear"
        .btnSaveBook.ControlTipText = "Save"
    End With

End Sub
