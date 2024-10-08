VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum st_operator
    ST_AND = 0
    ST_OR = 1
End Enum

Private Type objADODB
    cnn As ADODB.Connection
    rs As ADODB.Recordset
End Type

Private self As objADODB

Private Sub Class_Initialize()

    Set self.cnn = New ADODB.Connection
            
End Sub

Function connect()

    Dim strCnn As String
    Dim provider As String
    Dim dataSource As String
    
    provider = Environ$("PROVIDER")
    dataSource = Environ$("DATA_SOURCE")
    
    If provider = Empty Or dataSource = Empty Then
        Err.Raise 80000, Description:="Provider o data source not found"
    End If
    
'    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\books.accdb"
    strCnn = "Provider=" & provider & ";Data Source=" & ThisWorkbook.Path & dataSource
    
    On Error GoTo Catch
    
    With self.cnn
        .CommandTimeout = 30
        .ConnectionString = strCnn
        .Open
    End With
    
    Debug.Print "Connection open with database."
    
    Exit Function
    
Catch:
    handleError
End Function
Function disconnect()

    With self
        If Not .cnn Is Nothing Then
            With .cnn
                If .State <> adStateClosed Then
                    .Close
                End If
            End With
        End If
    End With
        
    With self
        If Not .rs Is Nothing Then
            With .rs
                If .State <> adStateClosed Then
                    .Close
                End If
            End With
        End If
    End With
        
    With self
        Set .cnn = Nothing
        Set .rs = Nothing
    End With
    
    Debug.Print "Connection closed with database."
    
End Function
Function readAll(tableName) As ADODB.Recordset
    
    Dim sql As String

    sql = "SELECT * FROM " & tableName
    
    On Error GoTo Catch
    
    With self
        .cnn.CursorLocation = adUseClient
        Set readAll = .cnn.Execute(sql)
    End With
    
    Exit Function
    
Catch:
    handleError
End Function
Function delete(tableName As String, filterParams As Dictionary)

    Dim sql As String
    Dim conditions As String
    
    On Error GoTo Catch
    
    conditions = generateConditionSQL(filterParams, ST_AND)
    conditions = " WHERE " & conditions
    
    sql = "SELECT * FROM " & tableName & conditions
    Set self.rs = New ADODB.Recordset

    With self.rs
        .CursorLocation = adUseClient
        .Open sql, self.cnn, adOpenDynamic, adLockOptimistic
        
        If Not (.BOF) And Not (.EOF) Then
            .MoveFirst
            Do While Not (.EOF)
                .delete
                .MoveNext
                Debug.Print "Record deleted"
            Loop
        Else
            Debug.Print "Records not found"
        End If
    End With
   
    Exit Function
    
Catch:
   handleError
End Function
Function read(tableName As String, filterParams As Dictionary) As ADODB.Recordset

    Dim sql As String
    Dim conditions As String
    
    On Error GoTo Catch
    conditions = generateConditionSQL(filterParams, ST_AND)
    sql = "SELECT * FROM " & tableName & " WHERE " & conditions

    Set self.rs = New ADODB.Recordset
    
    With self.rs
        .CursorLocation = adUseClient
        .Open sql, self.cnn ', adOpenStatic, adLockOptimistic
    End With
    
    Debug.Print "Get records with conditions."
    
    Set read = self.rs
    
    Exit Function
    
Catch:
    handleError
End Function
Function create(tableName As String, params As Dictionary)
    
    Dim sql As String
    
    On Error GoTo Catch
    
    sql = "SELECT * FROM " & tableName
    Set self.rs = New ADODB.Recordset
    
    With self.rs
        .CursorLocation = adUseClient
        .Open sql, self.cnn, adOpenStatic, adLockOptimistic
        
        .AddNew
        
        For Each k In params.Keys
            .fields(k).value = params(k)
        Next k
        
        .update
    End With
    
    Debug.Print "Insert new record!!!"
    Exit Function
    
Catch:
    handleError
End Function
Function update(tableName As String, filterParams As Dictionary, params As Dictionary)

    Dim sql As String
    Dim conditions As String
    
    On Error GoTo Catch
    
    conditions = generateConditionSQL(filterParams, ST_AND)
    conditions = " WHERE " & conditions
    
    sql = "SELECT * FROM " & tableName & conditions
    Set self.rs = New ADODB.Recordset

    With self.rs
        .CursorLocation = adUseClient
        .Open sql, self.cnn, adOpenKeyset, adLockOptimistic
        
        If Not (.BOF) And Not (.EOF) Then
            .MoveFirst
            Do While Not (.EOF)
                For Each k In params.Keys
                    .fields(k).value = params(k)
                Next k
                .update
                .MoveNext
            Loop
            
            Debug.Print "Update record!!!"
        Else
            Debug.Print "Records not found"
        End If
    End With
   
    Exit Function
    
Catch:
   handleError
End Function
Private Function generateConditionSQL(params As Dictionary, operator As st_operator) As String

    Dim conditions() As Variant
    Dim i As Integer
    Dim sql As String
    Dim op As String
    Const OP_AND As String = " AND "
    Const OP_OR As String = " OR "
    
    ReDim conditions(params.Count - 1)
    
    For Each k In params.Keys
        conditions(i) = k & " = " & convertToFieldValid(params(k))
        i = i + 1
    Next k
    
    Select Case operator
        Case Is = ST_AND
            op = OP_AND
        Case Is = ST_OR
            op = OP_OR
    End Select
    
    sql = Join(conditions, op)
    generateConditionSQL = sql
    
End Function
Private Function convertToFieldValid(value As Variant) As String
    
    Select Case True
        Case IsDate(value)
            convertToFieldValid = "#" & Format(value, "mm/dd/yyyy") & "#"
        Case IsNumeric(value)
            convertToFieldValid = value
        Case Else
            convertToFieldValid = "'" & value & "'"
    End Select
    
End Function
Private Function handleError()

    Err.Raise Err.Number, Description:=Err.Description
    On Error GoTo 0

End Function
