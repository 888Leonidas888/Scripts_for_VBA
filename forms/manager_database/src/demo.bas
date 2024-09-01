Attribute VB_Name = "demo"
Sub connect_and_disconnect()
    
    Dim st As New Storage
    
    With st
        .connect
        .disconnect
    End With
    
End Sub
Sub control_error()
    
    On Error GoTo Catch
    
    Dim st As New Storage
    
    With st
        .connect
        Err.Raise 9000, Description:="Error thrown"
        .disconnect
        Debug.Print "Everything OK"
    End With
    
    Exit Sub
    
Catch:
    Debug.Print "Error - number "; Err.Number
    Debug.Print Err.Description
    st.disconnect
    Debug.Print "Close for error"
End Sub
Sub insert_new_record()
    
    Dim st As New Storage
    Dim params As New Dictionary
    
    'create dictionary with fields and values to insert
    With params
        .Add "name_book", UCase("aplicaciones vba con excel")
        .Add "author", UCase("manuel torres remon")
        .Add "isbn", "978-60-762-2551-6"
        .Add "editorial", UCase("editorial macro")
        .Add "date_published", 2013
        .Add "badge", UCase("pen")
        .Add "price", 128.49
        .Add "created_at", #8/9/2024#
        .Add "updated_at", #8/9/2024#
    End With
    
    'call instance of storage for insert record
    With st
        .connect
        .create "books", params
        .disconnect
    End With
    
End Sub
Sub update_record()
    
    Dim st As New Storage
    Dim params As New Dictionary
    Dim filterParams As New Dictionary
    
    'create dictionary with fields and values to update
    With filterParams
        .Add "author", "franck ebel"
        .Add "editorial", "epsilon"
        .Add "created_at", #9/8/2024#
    End With
    
    With params
        .Add "badge", UCase("DOL")
    End With
    
    On Error GoTo Catch
    'call instance of storage for insert record
    With st
        .connect
        .update "books", filterParams, params
        .disconnect
    End With
    
    Exit Sub
    
    'control error
Catch:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, Title:="Error"
    Debug.Print Err.Number; " - "; Err.Description
    st.disconnect
End Sub
Sub read_records_with_conditions()
    
    Dim st As New Storage
    Dim filterParams As New Dictionary
    Dim rs As ADODB.Recordset
    
    'create dictionary with fields and values to insert record
    With filterParams
        .Add "updated_at", #8/9/2024#
'        .Add "name_book", UCase("curso de javascript")
    End With
    
    'Call instance of Storage for read
    With st
        .connect
        Set rs = .read("books", filterParams)
        
        'Read el recordset
        If Not (rs.BOF) And Not (rs.EOF) Then
            rs.MoveFirst
            Do While Not rs.EOF
                For i = 0 To rs.fields.Count - 1
                    Debug.Print rs.fields(i).value
                Next i
                Debug.Print " ********************* "
                rs.MoveNext
            Loop
        Else
            Debug.Print "Records not fount"
        End If
        
        .disconnect
    End With

End Sub
Sub delete_record()
    
    Dim st As New Storage
    Dim filterParams As New Dictionary

    With filterParams
        .Add "id", 26
    End With
    
    With st
        .connect
        .delete "books", filterParams
        .disconnect
    End With
End Sub
Sub read_all_records()
    
    Dim st As New Storage
    Dim rs As ADODB.Recordset
    
    With st
        .connect
        Set rs = .readAll("books")
        
        If Not (rs.BOF) And Not (rs.EOF) Then
            rs.MoveFirst
            Do While Not rs.EOF
                For i = 0 To rs.fields.Count - 1
                    Debug.Print rs.fields(i).value
                Next i
                Debug.Print " ********************* "
                rs.MoveNext
            Loop
        Else
            Debug.Print "Recordset Empty"
        End If
        
        .disconnect
    End With
    
End Sub

