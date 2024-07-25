# VBA IPv4

Esta biblioteca VBA proporciona una función para obtener la dirección IPv4 de la máquina.

## Contenido

- [Instalación](#instalación)
- [Función](#función)
  - [getIpv4](#getipv4)


## Instalación

1. Descarga el archivo `function_getIpv4.bas`.
2. Abre tu proyecto de VBA.
3. Importa el archivo `function_getIpv4.bas` a tu proyecto de VBA:
   - Abre el Editor de VBA (Alt + F11).
   - En la ventana del proyecto, haz clic derecho en tu proyecto y selecciona `Importar archivo...`.
   - Selecciona el archivo `function_getIpv4.bas` y haz clic en `Abrir`.
4. Activa las siguientes bibliotecas:
   - Microsoft Html Object Library
   - Microsoft VBScript Regular Expressions 5.5

### getIpv4

La función `getIpv4` obtiene la dirección IPv4 de la máquina.

#### Returns:

- `String`: La dirección IPv4 de la máquina.

#### Ejemplo:

```vb
sub test_ipv4()
    Dim myIp As String
    myIp = getIpv4()
    Debug.Print "Mi dirección IPv4 es: " & myIp
    'Mi dirección IPv4 es: 125.0.0.5
end sub
```
