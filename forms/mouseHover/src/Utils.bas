Attribute VB_Name = "Utils"
Sub formatForm(frm As UserForm)
    
    With frm
        .BackColor = vbWhite
        With .Frame1
            .BackColor = vbWhite
            .BorderStyle = 1
            .BorderColor = vbBlue
        End With
    End With
    
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.Label Then
            With ctrl
                .BackColor = vbWhite
'                .BorderStyle = 1
'                .BorderColor = vbBlue
            End With
        End If
    Next ctrl
    
End Sub
Sub chargeImageToLabel(frame As MSForms.frame)
    
    Dim pathImages As String
    Dim i As Byte
    
    i = 3
    pathImages = ThisWorkbook.Path & "\img\Imagen"
    For Each ctrl In frame.Controls
        If TypeOf ctrl Is MSForms.Label Then
            ctrl.Picture = LoadPicture(pathImages & i & ".jpg")
            i = i + 1
        End If
    Next ctrl
    
End Sub

