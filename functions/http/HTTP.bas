Attribute VB_Name = "HTTP"
'Este m�dulo le permite hacer peticiones HTTP de manera sencilla.

Public Type result
' Define un tipo personalizado llamado "result"
' para almacenar la respuesta de una solicitud HTTP.
    status As Integer
    responseText As String
    responseBody As Variant
End Type

Public Function HTTPRequests(endPoint As String, _
                            Optional queryParameters As String = "", _
                            Optional method As String = "GET", _
                            Optional body As String, _
                            Optional headers As Dictionary = Nothing) As result
    ' Realiza una solicitud HTTP y devuelve una estructura de tipo "result" con los detalles de la respuesta.
    '
    ' Args:
    '     endPoint (String): La URL del endpoint de la solicitud HTTP.
    '     queryParameters (Optional String): Par�metros de consulta para la URL, en formato de cadena. (Por defecto: "")
    '     method (Optional String): El m�todo HTTP a utilizar, como "GET", "POST", etc. (Por defecto: "GET")
    '     body (Optional String): El cuerpo de la solicitud, utilizado para m�todos como "POST".
    '     headers (Optional Dictionary): Un diccionario con los encabezados de la solicitud HTTP.
    '
    ' Returns:
    '     result: Una estructura que contiene el estado, el texto y el cuerpo de la respuesta HTTP.
    Dim request As New MSXML2.ServerXMLHTTP60
    Dim rsl As result
    Dim header As Variant
    
    With request
        .Open method, endPoint & queryParameters
        
        If Not headers Is Nothing Then
            For Each header In headers.Keys
                .setRequestHeader header, headers(header)
            Next header
        End If
        
        If LCase(methods) <> "get" Then
            .send body
        Else
            .send
        End If
        
        rsl.status = .status
        rsl.responseText = .responseText
        rsl.responseBody = .responseBody
    End With
    
    HTTPRequests = rsl
    
End Function

Public Function encodeURI(ByVal queryParameters As String) As String
    ' Codifica una cadena de par�metros de consulta para que sea compatible con una URL.
    '
    ' Args:
    '     queryParameters (String): La cadena de par�metros de consulta que necesita ser codificada.
    '
    ' Returns:
    '     String: La cadena de par�metros de consulta codificada.
    Dim characterSpecial As New Dictionary
    Dim key
    
    ' Crear un diccionario que mapea caracteres especiales a sus equivalentes en formato de porcentaje.
    ' El primer caracter de este diccionario debe ser porcentaje(%)
    With characterSpecial
        .Add "%", "%25"
        .Add " ", "%20"
        .Add "=", "%3D"
        .Add ",", "%2C"
        .Add """", "%22"
        .Add "<", "%3C"
        .Add ">", "%3E"
        .Add "#", "%23"
        .Add "|", "%7C"
        .Add "/", "%2F"
        .Add ":", "%3A"
        .Add "_", "%5F"
    End With
    
    For Each key In characterSpecial.Keys
        queryParameters = Replace(queryParameters, key, characterSpecial(key))
    Next key
    
    Set characterSpecial = Nothing
    
    encodeURI = queryParameters
    
End Function
