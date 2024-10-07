Attribute VB_Name = "connectionManager"
Function toSQLServer() As ADODB.Connection
    
    Dim config As New ADODB.Connection
    Dim server As String
    Dim database As String
    
    server = Environ("SERVER_SS")
    database = Environ("DDBB_SS")
    
    With config
        .provider = "SQLOLEDB"
        .Properties("Data Source") = server
        .Properties("Initial Catalog") = database
        .Properties("Integrated Security") = "SSPI"
    End With

    Set toSQLServer = config
    
End Function
Function toMySQL() As ADODB.Connection
    
    Dim config As New ADODB.Connection
    Dim server As String
    Dim port As String
    Dim database As String
    Dim user As String
    Dim pwd As String
    
    server = Environ("SERVER_MYSQL")
    port = Environ("PORT_MYSQL")
    database = Environ("DDBB_MYSQL")
    user = Environ("USER_MYSQL")
    pwd = Environ("MYSQL_PASS")
    
    With config
        .ConnectionString = "DRIVER={MySql ODBC 8.0 ANSI Driver};server=" _
                            & server & ";port=" & port & ";database=" & database
        .Properties("user id") = user
        .Properties("Password") = pwd
    End With

    Set toMySQL = config
    
End Function
Function toPostGreSQL() As ADODB.Connection
    
    Dim config As New ADODB.Connection
    Dim server As String
    Dim port As String
    Dim database As String
    Dim user As String
    Dim password As String
    
    server = Environ("SERVER_PG")
    port = Environ("PORT_PG")
    database = Environ("DDBB_PG")
    user = Environ("USER_PG")
    password = Environ("PWD_PG")
    
    With config
        .ConnectionString = "DRIVER={PostgreSQL ODBC Driver(ANSI)};server=" _
                            & server & ";port=" & port & ";database=" & database
        .Properties("user id") = user
        .Properties("password") = password
    End With

    Set toPostGreSQL = config
    
End Function
Function toMSAccess() As ADODB.Connection
    
    Dim config As New ADODB.Connection
    
    With config
        .provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Data Source") = ThisWorkbook.Path & "\db\books.accdb"
    End With

    Set toMSAccess = config
    
End Function
Function toMSExcel() As ADODB.Connection
    
    Dim config As New ADODB.Connection
    
    With config
        .provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Data Source") = ThisWorkbook.Path & "\db\books.xlsx"
        .Properties("Extended Properties") = "Excel 12.0 Xml;HDR=YES"
        
        ' libro con macros
'        .Properties("Extended Properties") = "Excel 12.0 Macro;HDR=YES"
    End With

    Set toMSExcel = config
    
End Function
