# Storage Class

La clase Storage proporciona una interfaz simple para interactuar con bases de datos utilizando **ADO (ActiveX Data Objects)** en **VBA (Visual Basic for Applications)**. Está diseñada para facilitar operaciones comunes de base de datos como conectar, desconectar, leer, actualizar, insertar y eliminar registros en tablas.

## Descripción

Esta clase maneja conexiones a bases de datos de Access y proporciona métodos para realizar operaciones **CRUD (Crear, Leer, Actualizar, Eliminar)**. Utiliza ADO para conectarse a bases de datos, ejecutar comandos SQL y gestionar recordsets.

## Instalación

### Importar la clase al proyecto VBA

Importa el siguiente archivo:

- [Storage.cls](./src/Storage.cls)

### Habilitar las referencias necesarias

- **Microsoft ActiveX Data Objects 6.1 Library**
- **Microsoft Scripting Runtime**
- **Microsoft VBScript Regular Expressions 5.5**

## Crea tus variables de entorno

Modifica las variables del fichero [env.bat](./env.bat) de ser necesario, estas variables serán leídas por la instancia de la **Clase Storage** para saber con que base de datos operar, antes de usar tu instancia debes ejecutar este archivo solo dándole doble click o ejecutando directamente estos comandos en tu terminal.

```sh
@REM Create variables enviroment
setx PROVIDER "Microsoft.ACE.OLEDB.12.0"
setx DATA_SOURCE "\db\books.accdb"
```
## Ejemplos

### Insertando un registro

El registro nuevo debe pasarse en un diccionario, el cual será pasado como segundo parámetro del método `create()` el primero deberá ser el nombre de la tabla.

```vb
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
```
Para ver mas ejemplos de uso, consulta el fichero [demo](./src/demo.bas).

>[!NOTE]
> Ten presente que las fechas se manejan en formato americano **mm/dd/yyyy**.