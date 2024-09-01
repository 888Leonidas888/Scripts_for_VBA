VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShowBooks 
   Caption         =   "UserForm1"
   ClientHeight    =   8640.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15405
   OleObjectBlob   =   "frmShowBooks.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmShowBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDeleteBook_Click()
    
    Dim rspta As Byte
    Dim id As Integer
    
    With lstShowBooks
        If .ListIndex = -1 Then
            MsgBox "Debe seleccionar un Item", vbExclamation, Title:="Antes de eliminar"
        Else
            rspta = MsgBox("Eliminar " & .List(.ListIndex, 0) & "?", _
            vbInformation + vbYesNo + vbDefaultButton2, "Eliminiar libro")
            
            id = .List(.ListIndex, 8)
            If rspta = vbYes Then
                Call data.deleteBookForID(id)
                Call data.showAllBooksInForms(lstShowBooks, lblRecordsFound)
            End If
        End If
    End With
    
End Sub

Private Sub btnEditBook_Click()

    Dim id As Integer
    Dim B As book
    
    With lstShowBooks
        If .ListIndex = -1 Then
            MsgBox "Debe seleccionar un Item", vbExclamation, Title:="Antes de editar"
        Else
            id = .List(.ListIndex, 8)
            Set B = data.getBook(id)
            Call data.editBookForm(B)
        End If
    End With
    
End Sub

Private Sub btnShowBooks_Click()

    Call data.showAllBooksInForms(lstShowBooks, lblRecordsFound)
    
End Sub

Private Sub btnAddBook_Click()

    With frmAddBook
        .btnSaveBook.Enabled = False
        .Show
    End With
    
End Sub

Private Sub UserForm_Initialize()

    With Me
        .Caption = "List books"
        .Frame1.Caption = Empty
    End With
    
    With Me
        .btnAddBook.ControlTipText = "Add"
        .btnDeleteBook.ControlTipText = "Delete"
        .btnEditBook.ControlTipText = "Edit"
        .btnShowBooks.ControlTipText = "Search"
    End With
    
    Call data.showAllBooksInForms(lstShowBooks, lblRecordsFound)
    
End Sub
