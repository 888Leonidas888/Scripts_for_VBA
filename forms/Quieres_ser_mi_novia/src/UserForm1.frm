VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14460
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const appName As String = Empty
Private Sub CommandButton1_Click()
    MsgBox "JaJaJaJa... sabia que dírias que sí", vbInformation, appName
    MsgBox "Tu respuesta ah sido aceptada....," & vbNewLine & "  Espera me cerraré en breve ...", vbInformation, appName
    Application.Wait Now + TimeValue("00:00:03")
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    MsgBox "No me lo esperaba", vbInformation, appName
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    With CommandButton2
        .Top = WorksheetFunction.RandBetween(78, 438)
        .Left = WorksheetFunction.RandBetween(0, 648)
    End With
    
End Sub

Private Sub UserForm_Initialize()

    CommandButton1.TabStop = False
    CommandButton2.TabStop = False
    
    With Me
        .Caption = "Elige una repuesta , por favor escoge con cuidado"
    End With
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        MsgBox "Espera no puedes salir sin antes darme tu respuesta", vbCritical, appName
        Cancel = True
    End If
End Sub


