VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Calendario"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblCalendar_Click()
    
    'Reemplazar por el control que recibirá la fecha,
    'esto debe hacerse antes que se muestre el calendario.
    Set ctrlFecha = Me.txtDate
    frmCalendar.Show
    
End Sub
