# VBA HTTP Request Utility

Esta biblioteca VBA proporciona funciones para realizar solicitudes HTTP y codificar URI.

## Contenido

- [Instalación](#instalación)
- [Funciones](#funciones)
  - [HTTPRequests](#httprequests)
  - [encodeURI](#encodeuri)
  - [createresource](#createresource)

## Instalación

1. Descarga el archivo `HTTPRequests.bas`.
2. Abre tu proyecto de VBA.
3. Importa el archivo `HTTPRequests.bas` a tu proyecto de VBA:
   - Abre el Editor de VBA (Alt + F11).
   - En la ventana del proyecto, haz clic derecho en tu proyecto y selecciona `Importar archivo...`.
   - Selecciona el archivo `HTTPRequests.bas` y haz clic en `Abrir`.
4. Activa las sgtes bibliotecas:
    - Microsoft Scripting Runtime
    - Microsoft XML, v6.0

### HTTPRequests

La función `HTTPRequests` realiza una solicitud HTTP y devuelve una estructura de tipo `result` con los detalles de la respuesta.

#### Args:

- `endPoint` (String): La URL del endpoint de la solicitud HTTP.
- `queryParameters` (Optional String): Parámetros de consulta para la URL, en formato de cadena. (Por defecto: "")
- `methods` (Optional String): El método HTTP a utilizar, como "GET", "POST", etc. (Por defecto: "GET")
- `body` (Optional String): El cuerpo de la solicitud, utilizado para métodos como "POST".
- `headers` (Optional Dictionary): Un diccionario con los encabezados de la solicitud HTTP.

#### Returns:

- `result`: Una estructura que contiene el estado, el texto y el cuerpo de la respuesta HTTP.

#### Ejemplo:

```vb
Dim headers As New Dictionary
headers.Add "Content-Type", "application/json"

Dim response As result
response = HTTPRequests("https://api.example.com/data", "", "GET", , headers)

Debug.Print "Status: " & response.status
Debug.Print "Response Text: " & response.responseText
```
### encodeURI
La función encodeURI codifica una cadena de parámetros de consulta para que sea compatible con una URL.

#### Args:
- `queryParameters` (String): La cadena de parámetros de consulta que necesita ser codificada.
#### Returns:
- `String`: La cadena de parámetros de consulta codificada.

#### Ejemplo:

```vb
Dim encodedParams As String
encodedParams = encodeURI("name=John Doe&age=30")

Debug.Print "Encoded Parameters: " & encodedParams
'Encoded Parameters: name%3DJohn%20Doe&age%3D30
```
### createResource

La función `createResource` crea un recurso a partir de un array de bytes y lo guarda en la ruta especificada.

#### Args:
- `resource` (Byte()): Array de bytes que representa el recurso a escribir en el archivo.
- `resourcePathDestination` (String): La ruta donde se guardará el archivo del recurso.

#### Returns:
- `Boolean`: Devuelve True si el recurso se creó y guardó exitosamente.

#### Raises:
- `Error 40000`: Si ocurre un error al escribir el recurso, se genera un error con el mensaje "Error writing resource".

#### Ejemplo:

```vb
Dim myResource() As Byte
' Rellenar el array de bytes con el recurso que deseas guardar
Dim success As Boolean

success = createResource(myResource, "C:\path\to\your\resource.jpg")

If success Then
    MsgBox "El recurso se guardó exitosamente."
Else
    MsgBox "Hubo un error al guardar el recurso."
End If
```