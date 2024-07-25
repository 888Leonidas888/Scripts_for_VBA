Attribute VB_Name = "QRGenerator"
Sub creat()
      
    Dim result As HTTP.result
    Dim url As String
    Dim resource() As Byte
    Dim resourcePathDestination As String
    
    On Error GoTo Catch
    
    resourcePathDestination = ThisWorkbook.Path & "\miQR.png"
    
    'Proveedores para c�digo QR
    'url = "https://qrickit.com/api/qr.php?d=https://google.com&addtext=google&txtcolor=442EFF&fgdcolor=76103C&bgdcolor=C0F912&qrsize=150&t=p&e=m"
    'url = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=Hello+world"
    '----
    result = HTTP.HTTPRequests(url)

    With result
        If .status = 200 Then
            resource = .responseBody
             If createResource(resource, resourcePathDestination) Then
                Debug.Print "Qr creado"
             End If
        End If
    End With
    
    Exit Sub
    
Catch:
    Debug.Print Err.Number
    Debug.Print Err.Description
    
End Sub
