VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegister 
   Caption         =   "Registro"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "frmRegister.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If utils.isInputsEmpty(Me.Frame1) Then
        MsgBox "Complete el campo antes de continuar", vbExclamation
    Else
        MsgBox "Bien, Todos los datos han sido ingresados", vbInformation
    End If
End Sub

Private Sub CommandButton2_Click()
    Call utils.clearInputs(Me.Frame1)
End Sub

Private Sub UserForm_Initialize()
    With ComboBox1
        .AddItem "und"
        .AddItem "pkg"
    End With
End Sub
