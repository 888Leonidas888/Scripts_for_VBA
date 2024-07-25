# VBA QR Code Generator

Esta biblioteca VBA proporciona una función para crear códigos QR utilizando una API y guardar la imagen resultante en el disco.

## Contenido

- [Instalación](#instalación)
- [Funciones](#funciones)
  - [create_Qr](#create_qr)
- [Uso](#Uso)
- [Proveedores para el API](#proveedores-para-el-api)

## Instalación

1. Descarga los archivos [HTTP.bas](../http/) y `QRGenerator.bas`.
2. Abre tu proyecto de VBA.
3. Importa los archivos a tu proyecto de VBA:
   - Abre el Editor de VBA (Alt + F11).
   - En la ventana del proyecto, haz clic derecho en tu proyecto y selecciona `Importar archivo...`.
   - Selecciona los archivos [HTTP.bas](../http/) y `QRGenerator.bas` y haz clic en `Abrir`.

## Uso

Para este caso se usan 2 proveedores para generar el QR, descomenta el contenido de la variable `url` para probar con cada proveedor.

```vb
Sub creat()

    Dim result As HTTP.result
    Dim url As String
    Dim resource() As Byte
    Dim resourcePathDestination As String

    On Error GoTo Catch

    resourcePathDestination = ThisWorkbook.Path & "\miQR.png"

    'Proveedores para código QR
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

```

## Proveedores para el API

Para saber mas sobre el uso de cada API visite la documentación de cada proveedor:

- [QRickit QR Code API](https://qrickit.com/qrickit_apps/qrickit_api.php)
- [QR code](https://goqr.me/api/doc/)

Este proveedor solicita suscripción para probar su api.

- [QRCodeMonkey](https://www.qrcode-monkey.com/qr-code-api-with-logo/)
