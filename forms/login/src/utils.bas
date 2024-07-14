Attribute VB_Name = "utils"
Sub settingLogin(frm As UserForm)

    Dim imagePath As String
    imagePath = ThisWorkbook.Path & "\img\city.jpg"
    
    With frm
        If Len(Dir(imagePath)) > 0 Then
            With .imgLogin
                .Picture = LoadPicture(imagePath)
                .PictureSizeMode = fmPictureSizeModeStretch
            End With
        End If
        
        With .txtPassword
            .PasswordChar = "*"
        End With
        
        .btnLogIn.Default = True
        
    End With
    
End Sub
Sub styleLogin(frm As UserForm)
    
    Const hexColorMain As String = "#DCF1EB"
    
    
    frm.BackColor = converterRGBForVBA(hexColorMain)

    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.Label Then
            ctrl.BackStyle = 0
        ElseIf TypeOf ctrl Is MSForms.TextBox Then
            With ctrl
                .BackColor = converterRGBForVBA(hexColorMain)
                .BorderStyle = fmBorderStyleSingle
                .SpecialEffect = fmSpecialEffectFlat
                .BorderColor = converterRGBForVBA("#38D9A9")
            End With
        ElseIf TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.BackColor = converterRGBForVBA(hexColorMain)
        ElseIf TypeOf ctrl Is MSForms.Frame Then
            ctrl.BackColor = converterRGBForVBA("#88DBC2")
        ElseIf TypeOf ctrl Is MSForms.CommandButton Then
            With ctrl
                .BackColor = converterRGBForVBA("#38D9A9")
                .ForeColor = vbWhite
            End With
        End If
    Next ctrl
    
End Sub
Public Function converterRGBForVBA(hexColor As String) As Variant
    
    Dim R As String
    Dim G As String
    Dim B As String
    
    If Len(hexColor) <> 7 Then: Err.Raise Number:=2000, Description:="Len invalid"
    If Left(hexColor, 1) <> "#" Then: Err.Raise Number:=2001, Description:="Firts character invalid"
    
    hexColor = Right(hexColor, Len(hexColor) - 1)
    
    R = Left(hexColor, 2)
    G = Mid(hexColor, 3, 2)
    B = Right(hexColor, 2)
    
    converterRGBForVBA = "&H" & B & G & R
    
End Function

