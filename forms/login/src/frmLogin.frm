VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Inicio de sesión"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogIn_Click()
    
    Dim lblWelcome As MSForms.Label
    
    If Trim(txtUser) <> Empty And Trim(txtPassword) <> Empty Then
        If storage.validateLogin(txtUser, txtPassword) Then
            Set lblWelcome = frmMain.Label1
            
            lblWelcome = "Bienvenido " & txtUser
            
            Unload Me
            
            'reemplazar por el formulario a abrir.
            frmMain.Show
            '------------------------------------
        Else
            MsgBox "Sus credenciales son incorrectas", vbExclamation
        End If
    Else
        MsgBox "Debe completar el usuario y contraseña", vbExclamation
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Call utils.settingLogin(Me)
    Call utils.styleLogin(Me)
    
End Sub
