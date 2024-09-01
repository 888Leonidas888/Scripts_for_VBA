VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddBook 
   Caption         =   "UserForm1"
   ClientHeight    =   8985.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   OleObjectBlob   =   "frmAddBook.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAddBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddBook_Click()

    Dim newBook As New book
        
    On Error GoTo Catch
    
    With newBook
        .name = txtName
        .author = txtAuthor
        .isbn = txtISBN
        .comment = txtComment
        .editorial = txtEditorial
        .published = txtPublished
        .badge = cmbBadge
        .created = txtCreated
        .updated = txtUpdate
        .price = txtPrice
    End With
    
    Call data.insertNewBook(newBook)
    MsgBox "Nuevo libro: " & txtName & " registrado", vbInformation
    
    Exit Sub
    
Catch:
    MsgBox Err.Number & vbCrLf & Err.Description
    On Error GoTo 0
End Sub

Private Sub btnClear_Click()
    Call utils.clearInputs(Me.Frame1)
    Call utils.settingFrmAddBooks(Me)
End Sub

Private Sub btnSaveBook_Click()

    Dim editBook As New book
    Dim frm As UserForm
            
    Set frm = frmShowBooks
    
    With editBook
        .id = Me.Frame1.Caption
        .name = txtName
        .author = txtAuthor
        .isbn = txtISBN
        .comment = txtComment
        .editorial = txtEditorial
        .published = txtPublished
        .badge = cmbBadge
        .created = txtCreated
        .updated = txtUpdate
        .price = txtPrice
    End With
    
    Call data.updateBook(editBook)
    Call data.showAllBooksInForms(frm.lstShowBooks, frm.lblRecordsFound)
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Call settingFrmAddBooks(Me)
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Dim frm As UserForm
    
    Set frm = frmShowBooks
    
    Call data.showAllBooksInForms(frm.lstShowBooks, frm.lblRecordsFound)
    
End Sub
