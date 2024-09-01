# CRUD with VBA and SQL using Access

Este archivo tiene un formulario de ejemplo de como utilizar la [clase Storage](../manager_database/src/Storage.cls) para hacer operaciones **CRUD**.

# Instalación

### Habilitar las siguientes referencias:

- **Microsoft ActiveX Data Objects 6.1 Library**
- **Microsoft Scripting Runtime**
- **Microsoft VBScript Regular Expressions 5.5**

### Cree sus varaibles de entorno

Ejecute los siguientes comandos en su terminal o ejecute el archivo [env.bat](./env.bat), estas variables serán leídas por la instancia de [clase storage](./src/Storage.cls).

```sh
@REM Create variables enviroment
setx PROVIDER "Microsoft.ACE.OLEDB.12.0"
setx DATA_SOURCE "\db\books.accdb"
```
# Ejecución

Abra el archivo [storage.xlsm](./storage.xlsm), habilite las macros para dar permisos, presione el botón **Open app** esto le mostrará un formulario, ahora esta listo para probarlo.

# Enlaces

- [iconos usados en este proyecto](https://icon-icons.com/pack/Windows-8-Metro-Icons/17)