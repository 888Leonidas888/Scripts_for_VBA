VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Book sale"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    Dim lstDetails As MSForms.ListBox
    
    Set lstDetails = frmSecond.ListBox1
    
    With lstDetails
        .ColumnCount = 2
        .ColumnWidths = 60
        .AddItem "Name : "
        .List(.ListCount - 1, 1) = Me.TextBox1.Text
        
        .AddItem "Author : "
        .List(.ListCount - 1, 1) = Me.TextBox2.Text
        
        .AddItem "ISBN : "
        .List(.ListCount - 1, 1) = Me.TextBox3.Text
        
        .AddItem "Price : "
        .List(.ListCount - 1, 1) = Format(Me.TextBox4.Text, "#,##0.00$")
    End With

    frmSecond.Show
    
End Sub

Private Sub UserForm_Initialize()
    
    TextBox1.Value = "Curso de JavaScript"
    TextBox2.Value = "Astor de Caso Parra"
    TextBox3.Value = "978-84-415-4228-0"
    TextBox4.Value = Format(85.95, "#,###0.00")
    
End Sub
