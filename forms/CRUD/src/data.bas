Attribute VB_Name = "data"
Sub handleError(st As Storage)
    MsgBox Err.Number & vbCrLf & Err.Description
    st.disconnect
End Sub

Sub insertNewBook(book As book)
    
    Dim st As New Storage
    Dim params As New Dictionary
    
    On Error GoTo Catch
    'create dictionary with fields and values to insert
    With params
        .Add "name_book", book.name
        .Add "author", book.author
        .Add "isbn", book.isbn
        .Add "editorial", book.editorial
        .Add "date_published", book.published
        .Add "badge", book.badge
        .Add "price", book.price
        .Add "created_at", book.created
        .Add "updated_at", book.updated
        .Add "comment", book.comment
    End With
    
'    call instance of storage for insert record
    With st
        .connect
        .create "books", params
        .disconnect
    End With
    
    Exit Sub
      
Catch:
    Call handleError(st)
End Sub
Function getAllBooks() As Collection
    
    Dim st As New Storage
    Dim rs As ADODB.Recordset
    Dim B As book
    Dim books As New Collection
    Dim comment As Variant
    
    On Error GoTo Catch
    
    With st
        .connect
        Set rs = .readAll("books")
        
        If Not (rs.BOF) And Not (rs.EOF) Then
            rs.MoveFirst
            Do While Not rs.EOF
                Set B = New book
                
                B.id = rs.fields("id").value
                B.author = rs.fields("author").value
                B.name = rs.fields("name_book").value
                B.isbn = rs.fields("isbn").value
                B.editorial = rs.fields("editorial").value
                B.published = rs.fields("date_published").value
                B.badge = rs.fields("badge").value
                B.price = rs.fields("price").value
                B.created = rs.fields("created_at").value
                B.updated = rs.fields("updated_at").value
                
                comment = rs.fields("comment").value
                B.comment = IIf(IsNull(comment), Empty, comment)
                
                rs.MoveNext
                
                books.Add B
            Loop
        End If
        .disconnect
    End With
    
    If Not books Is Nothing Then
        Set getAllBooks = books
    Else
        Set getAllBooks = books
    End If
    
    Exit Function
    
Catch:
    handleError (st)
End Function
Sub deleteBookForID(id As Integer)
    
    Dim st As New Storage
    Dim filterParams As New Dictionary

    With filterParams
        .Add "id", id
    End With
    
    On Error GoTo Catch
    
    With st
        .connect
        .delete "books", filterParams
        .disconnect
    End With
    
    Exit Sub
Catch:
    Call handleError(st)
End Sub
Function getBook(id As Integer) As book

    Dim st As New Storage
    Dim filterParams As New Dictionary
    Dim rs As ADODB.Recordset
    Dim B As New book
    
    With filterParams
        .Add "id", id
    End With
    
    On Error GoTo Catch
    
    With st
        .connect
        Set rs = .read("books", filterParams)

        If Not (rs.BOF) And Not (rs.EOF) Then
            B.id = rs.fields("id").value
            B.author = rs.fields("author").value
            B.name = rs.fields("name_book").value
            B.isbn = rs.fields("isbn").value
            B.editorial = rs.fields("editorial").value
            B.published = rs.fields("date_published").value
            B.badge = rs.fields("badge").value
            B.price = rs.fields("price").value
            B.created = rs.fields("created_at").value
            B.updated = rs.fields("updated_at").value
            
            comment = rs.fields("comment").value
            B.comment = IIf(IsNull(comment), Empty, comment)
        End If
        .disconnect
    End With
    
    If Not B Is Nothing Then
        Set getBook = B
    Else
        Set getBook = Nothing
    End If
    
    Exit Function
Catch:
    handleError (st)
End Function
Sub updateBook(book As book)
    
    Dim st As New Storage
    Dim params As New Dictionary
    Dim filterParams As New Dictionary
    
    'create dictionary with fields and values to update
    With filterParams
        .Add "id", book.id
    End With
    
    With params
        .Add "name_book", book.name
        .Add "author", book.author
        .Add "isbn", book.isbn
        .Add "editorial", book.editorial
        .Add "date_published", book.published
        .Add "badge", book.badge
        .Add "price", book.price
        .Add "created_at", book.created
        .Add "updated_at", book.updated
        .Add "comment", book.comment
    End With
    
    On Error GoTo Catch
    
    With st
        .connect
        .update "books", filterParams, params
        .disconnect
    End With
    
    Exit Sub
    
Catch:
   handleError (st)
End Sub
Sub editBookForm(book As book)
    
    With frmAddBook
        .txtAuthor = book.author
        .txtComment = book.comment
        .txtCreated = book.created
        .txtEditorial = book.editorial
        .txtISBN = book.isbn
        .txtName = book.name
        .txtPrice = Format(book.price, "#,##0.00")
        .txtPublished = book.published
        .txtUpdate = book.updated
        .cmbBadge = book.badge
        .Caption = "Estas editando un libro"
        .Frame1.Caption = book.id
        .btnAddBook.Enabled = False
        .Show
    End With
    
End Sub
Sub showAllBooksInForms(lst As MSForms.ListBox, lbl As MSForms.Label)

    Dim books As Collection
    Dim B As New book
    
    Set books = data.getAllBooks()
    
    If Not books Is Nothing Then
        With lst
            .Clear
            .ColumnCount = 9
            .ColumnWidths = "140;130;100;100;100;30;50;50;0"
            .ListStyle = fmListStyleOption
            For Each B In books
                .AddItem B.name
                .List(.ListCount - 1, 1) = B.author
                .List(.ListCount - 1, 2) = B.isbn
                .List(.ListCount - 1, 3) = B.comment
                .List(.ListCount - 1, 4) = B.editorial
                .List(.ListCount - 1, 5) = B.badge
                .List(.ListCount - 1, 6) = B.published
                .List(.ListCount - 1, 7) = Format(B.price, "#,##0.00")
                .List(.ListCount - 1, 8) = B.id
            Next B
        End With
    End If
    
    With lbl
        .Font.Bold = True
        .Caption = "Books found " & lst.ListCount
    End With
End Sub


