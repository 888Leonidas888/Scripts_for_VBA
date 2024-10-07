Attribute VB_Name = "demo"
Sub connect_with_config_any()

    Dim st As New Storage
 
    On Error GoTo Catch
    
    With st
        .connect connectionManager.toSQLServer
        .disconnect
    End With
    
    Exit Sub
    
Catch:

    With Err
        Debug.Print .Number
        Debug.Print .Description
    End With
    
    st.disconnect
    
End Sub

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
        .Add "name_book", UCase("sql los fundamentos del lenguaje")
        .Add "author", UCase("anne christine bisson")
        .Add "isbn", "978-2-409-03037-6"
        .Add "editorial", UCase("eni")
        .Add "date_published", 2019
        .Add "badge", UCase("pen")
        .Add "price", 137.12
        .Add "created_at", #10/6/2024#
        .Add "updated_at", #10/6/2024#
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
        .Add "id", 29
    End With
    
    With params
        .Add "price", 159.72
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
        .Add "id", 28
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

