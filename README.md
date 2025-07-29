# Procesador de Excel a SQL Server

Aplicación de escritorio construida con Electron para limpiar archivos Excel y enviar los datos a una base de datos SQL Server. (with actions)

## Características

- Selección de archivos Excel (`.xlsx`, `.xls`).
- Limpieza y validación de datos según campos requeridos.
- Exportación de un archivo Excel limpio.
- Envío masivo de datos limpios a una tabla SQL Server.
- Interfaz sencilla y mensajes de estado en pantalla.

## Instalación

1. Clona este repositorio.
2. Instala las dependencias:

   ```sh
   npm install
   ```

3. Configura tus credenciales de SQL Server en el archivo [`mainori.js`](mainori.js):

   ```js
   // ...existing code...
   const dbConfig = {
       server: 'TU_SERVIDOR_SQL',
       authentication: {
           type: 'default',
           options: {
               userName: 'TU_USUARIO',
               password: 'TU_CONTRASEÑA'
           }
       },
       options: {
           encrypt: true,
           database: 'TU_BASE_DE_DATOS',
           trustServerCertificate: true
       }
   };
   // ...existing code...
   ```

## Uso

1. Inicia la aplicación:

   ```sh
   npm start
   ```

2. Haz clic en **"1. Generar Tabular Limpio"** para seleccionar y limpiar un archivo Excel.
3. Haz clic en **"2. Enviar a Base de Datos"** para cargar los datos limpios a SQL Server.

## Estructura del Proyecto

- [`mainori.js`](mainori.js): Lógica principal de Electron y procesamiento de archivos.
- [`preload.js`](preload.js): API segura para comunicación entre frontend y backend.
- [`tabular_limpio.js`](tabular_limpio.js): Lógica de la interfaz y manejo de eventos.
- [`index.html`](index.html), [`index.css`](index.css): Interfaz de usuario.

## Dependencias principales

- [Electron](https://www.electronjs.org/)
- [ExcelJS](https://github.com/exceljs/exceljs)
- [Tedious](https://github.com/tediousjs/tedious)
- [Lodash](https://lodash.com/)

## Notas

- Asegúrate de tener acceso a la base de datos SQL Server y de que la tabla `PackingItems` existe con las columnas requeridas.
- Para producción, utiliza variables de entorno para las credenciales.

---

Desarrollado por Jandrey.
