const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('node:path');
const fs = require('fs').promises;
const ExcelJS = require('exceljs');
// ---> INICIO: Nuevos 'requires' para la base de datos y el sistema
const { Connection, Request, TYPES } = require('tedious');
const os = require('os');
// ---> FIN: Nuevos 'requires'

// Constantes para configuración
const CONFIG = {
    ROWS_TO_SKIP: 3,
    REQUIRED_FIELDS: ['item', 'n_partida', 'nombre_del_articulo'],
    EXCEL_EXTENSIONS: ['xlsx', 'xls'],
    OUTPUT_FILENAME: 'PACKING_limpio.xlsx'
};

// ---> INICIO: Configuración de la Base de Datos
// ¡MUY IMPORTANTE! Reemplaza estos valores con tus credenciales reales.
// Para mayor seguridad en producción, considera usar variables de entorno o un sistema de gestión de secretos.
const dbConfig = {
    server: 'TU_SERVIDOR_SQL', // Ejemplo: 'localhost' o 'servidor.database.windows.net'
    authentication: {
        type: 'default',
        options: {
            userName: 'TU_USUARIO',
            password: 'TU_CONTRASEÑA'
        }
    },
    options: {
        encrypt: true, // Requerido para Azure SQL, puede ser false para servidores locales sin SSL.
        database: 'TU_BASE_DE_DATOS',
        trustServerCertificate: true // Poner en 'false' en producción si tienes un certificado de servidor válido.
    }
};
// ---> FIN: Configuración de la Base de Datos


// ========================================================================
// === INICIO: LÓGICA ORIGINAL DE PROCESAMIENTO DE EXCEL (SIN CAMBIOS) ====
// ========================================================================

/**
 * Valida y limpia los datos de una fila de Excel
 * @param {Object} row - Fila de Excel
 * @param {number} rowNumber - Número de fila para mensajes de error
 * @returns {Object|null} Datos limpios o null si la fila no es válida
 */
function processExcelRow(row, rowNumber) {
    const getCellValue = (cellIndex, type = 'string', required = false) => {
        const cell = row.getCell(cellIndex);
        if (!cell || cell.value === null || cell.value === undefined || cell.value === '') {
            if (required) {
                throw new Error(`Celda requerida vacía en la fila ${rowNumber}, columna ${cellIndex}`);
            }
            return type === 'number' ? null : '';
        }

        const value = String(cell.value).trim();
        
        if (type === 'number') {
            const numValue = parseFloat(value.replace(',', '.'));
            return isNaN(numValue) ? null : numValue;
        }
        
        return value;
    };

    try {
        const rowData = {
            item: getCellValue(1, 'number', true),
            n_partida: getCellValue(2, 'string', true),
            nombre_del_articulo: getCellValue(5, 'string', true),
            articulo_descripcion: getCellValue(6, 'string'),
            colores: getCellValue(10, 'string'),
            cod_colores: getCellValue(13, 'number'),
            pqt: getCellValue(16, 'number'),
            kg: getCellValue(18, 'number')
        };

        const missingFields = CONFIG.REQUIRED_FIELDS.filter(field => 
            rowData[field] === null || rowData[field] === '' || rowData[field] === undefined
        );

        if (missingFields.length > 0) {
            console.warn(`Fila ${rowNumber}: Campos requeridos faltantes: ${missingFields.join(', ')}`);
            return null;
        }

        return rowData;
    } catch (error) {
        console.error(`Error procesando fila ${rowNumber}:`, error.message);
        return null;
    }
}

/**
 * Lee y procesa un archivo Excel
 * @param {string} filePath - Ruta al archivo Excel
 * @returns {Promise<Array>} Datos procesados
 */
async function readExcelFileMain(filePath) {
    try {
        await fs.access(filePath);
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
            throw new Error('El archivo Excel no contiene hojas de trabajo');
        }

        const jsonData = [];
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber <= CONFIG.ROWS_TO_SKIP) return;
            const rowData = processExcelRow(row, rowNumber);
            if (rowData) jsonData.push(rowData);
        });

        console.log(`Procesamiento completado: ${jsonData.length} filas válidas encontradas.`);
        return jsonData;

    } catch (error) {
        console.error('Error al leer el archivo Excel:', error.message);
        throw new Error(`No se pudo procesar el archivo: ${error.message}`);
    }
}

/**
 * Exporta los datos procesados a un archivo Excel
 * @param {Array} jsonData - Datos a exportar
 * @param {string} outputPath - Ruta de salida del archivo
 * @returns {Promise<string>} Ruta del archivo generado
 */
async function exportCleanedDataToExcelMain(jsonData, outputPath) {
    // ... (Tu función de exportación es excelente, la dejo exactamente igual)
    if (!Array.isArray(jsonData) || jsonData.length === 0) {
        throw new Error('No hay datos válidos para exportar');
    }
    console.log(`\nIniciando exportación a: ${outputPath}`);
    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Excel Processor App';
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet('Datos Limpios', { pageSetup: { paperSize: 9, orientation: 'landscape' } });
        worksheet.columns = [
            { header: 'Item', key: 'item', width: 15, style: { numFmt: '0' } },
            { header: 'No. Partida', key: 'n_partida', width: 20 },
            { header: 'Nombre del Artículo', key: 'nombre_del_articulo', width: 40 },
            { header: 'Descripción del Artículo', key: 'articulo_descripcion', width: 50 },
            { header: 'Colores', key: 'colores', width: 25 },
            { header: 'Cod. Colores', key: 'cod_colores', width: 15, style: { numFmt: '0' } },
            { header: 'Paquete (PQT)', key: 'pqt', width: 15, style: { numFmt: '#,##0' } },
            { header: 'Kilogramos (KG)', key: 'kg', width: 15, style: { numFmt: '#,##0.00' } }
        ];
        worksheet.addRows(jsonData);
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } };
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                const columnLength = cell.value ? cell.value.toString().length : 0;
                maxLength = Math.max(maxLength, columnLength);
            });
            column.width = Math.min(Math.max(column.width || 0, maxLength + 2), 50);
        });
        await worksheet.protect('', { selectLockedCells: false, selectUnlockedCells: true });
        const dir = path.dirname(outputPath);
        await fs.mkdir(dir, { recursive: true });
        await workbook.xlsx.writeFile(outputPath);
        console.log(`Archivo exportado exitosamente: ${outputPath}`);
        return outputPath;
    } catch (error) {
        console.error('Error al exportar el archivo Excel:', error.message);
        throw new Error(`Error al guardar el archivo: ${error.message}`);
    }
}

// ======================================================================
// === FIN: LÓGICA ORIGINAL DE PROCESAMIENTO DE EXCEL (SIN CAMBIOS) ====
// ======================================================================


/**
 * Crea la ventana principal de la aplicación
 */
function createWindow() {
    // ... (Tu función createWindow es perfecta, no necesita cambios)
    const mainWindow = new BrowserWindow({
        width: 1280, height: 800, minWidth: 800, minHeight: 600, show: false,
        webPreferences: {
            // Asegúrate de que el nombre del preload script coincida con el que creamos
            preload: path.join(__dirname, 'preload.js'), // Nombre estandarizado
            nodeIntegration: false, contextIsolation: true, sandbox: true, enableRemoteModule: false
        },
        icon: path.join(__dirname, 'assets/icon.png')
    });
    mainWindow.loadFile('index.html');
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
        if (process.env.NODE_ENV === 'development') mainWindow.webContents.openDevTools();
    });
    mainWindow.on('closed', () => { mainWindow.destroy(); });
    return mainWindow;
}


// ======================================================================
// === INICIO: MANEJADORES DE IPC (EXISTENTES Y NUEVOS) ==================
// ======================================================================

/**
 * Manejador para el diálogo de apertura de archivo (EXISTENTE)
 */
ipcMain.handle('open-file-dialog', async () => {
    try {
        const { canceled, filePaths } = await dialog.showOpenDialog({
            title: 'Seleccionar archivo Excel', properties: ['openFile'],
            filters: [
                { name: 'Archivos Excel', extensions: CONFIG.EXCEL_EXTENSIONS },
                { name: 'Todos los archivos', extensions: ['*'] }
            ]
        });
        if (!canceled && filePaths.length > 0) {
            const stats = await fs.stat(filePaths[0]);
            if (stats.size === 0) throw new Error('El archivo está vacío');
            return filePaths[0];
        }
        return null;
    } catch (error) {
        console.error('Error en el diálogo de apertura:', error);
        throw new Error(`No se pudo abrir el archivo: ${error.message}`);
    }
});

/**
 * Manejador para procesar y exportar el archivo Excel (EXISTENTE)
 */
ipcMain.handle('process-excel-and-export', async (event, filePath) => {
    try {
        if (!filePath) throw new Error('No se proporcionó una ruta de archivo válida');
        const stats = await fs.stat(filePath);
        if (stats.size === 0) throw new Error('El archivo está vacío');
        const cleanedData = await readExcelFileMain(filePath);
        if (!cleanedData || cleanedData.length === 0) throw new Error('No se encontraron datos válidos para exportar');
        const outputPath = path.join(app.getPath('documents'), `${path.parse(filePath).name}_${CONFIG.OUTPUT_FILENAME}`);
        const savedPath = await exportCleanedDataToExcelMain(cleanedData, outputPath);
        return { success: true, outputPath: savedPath, processedRows: cleanedData.length };
    } catch (error) {
        console.error('Error al procesar el archivo:', error);
        return { success: false, error: error.message, stack: process.env.NODE_ENV === 'development' ? error.stack : undefined };
    }
});


// ---> INICIO: NUEVOS MANEJADORES Y LÓGICA DE BASE DE DATOS <---

/**
 * Manejador NUEVO para leer un archivo Excel y devolver los datos en formato JSON.
 * Se usará para preparar los datos antes de enviarlos a la BD.
 */
ipcMain.handle('read-cleaned-excel', async (event, filePath) => {
    try {
        if (!filePath) throw new Error('No se proporcionó una ruta de archivo.');
        // Reutilizamos tu función de lectura y limpieza. ¡Perfecto!
        const jsonData = await readExcelFileMain(filePath);
        if (!jsonData || jsonData.length === 0) {
            return { success: false, error: 'El archivo no contiene datos válidos después de la limpieza.' };
        }
        return { success: true, data: jsonData };
    } catch (error) {
        console.error('Error al leer el archivo para la BD:', error);
        return { success: false, error: error.message };
    }
});

/**
 * Manejador NUEVO para enviar los datos a la base de datos SQL Server.
 */
ipcMain.handle('send-data-to-db', async (event, data) => {
    if (!data || data.length === 0) {
        return { success: false, error: 'No hay datos para enviar a la base de datos.' };
    }

    // Usamos una promesa para manejar la naturaleza asíncrona de la conexión a la BD.
    return new Promise((resolve) => {
        const connection = new Connection(dbConfig);

        connection.on('connect', (err) => {
            if (err) {
                console.error('Error de conexión a la BD:', err);
                return resolve({ success: false, error: `Error de conexión: ${err.message}` });
            }
            console.log('Conectado a la base de datos. Iniciando inserción masiva...');
            executeBulkInsert(connection, data, resolve);
        });

        connection.on('error', (err) => {
            console.error('Error general de la conexión a la BD:', err);
            // Evita resolver la promesa si ya se cerró por otro error.
            if (!connection.closed) {
                resolve({ success: false, error: `Error de red o conexión: ${err.message}` });
            }
        });
        
        connection.connect();
    });
});

/**
 * Función auxiliar para ejecutar una inserción masiva (Bulk Insert) con Tedious.
 * @param {Connection} connection - La instancia de conexión activa.
 * @param {Array} data - Los datos a insertar.
 * @param {Function} resolve - La función para resolver la promesa del manejador IPC.
 */
function executeBulkInsert(connection, data, resolve) {
    // Asegúrate de que el nombre de la tabla ('PackingItems') y las columnas coinciden con tu base de datos.
    const request = new Request(
        `INSERT INTO PackingItems (item, n_partida, nombre_del_articulo, articulo_descripcion, colores, cod_colores, pqt, kg) 
         VALUES (@item, @n_partida, @nombre_del_articulo, @articulo_descripcion, @colores, @cod_colores, @pqt, @kg)`,
        (err, rowCount) => {
            connection.close(); // Cierra la conexión al terminar o en caso de error.
            if (err) {
                console.error('Error en Bulk Insert:', err);
                resolve({ success: false, error: `Error al insertar datos: ${err.message}` });
            } else {
                console.log(`${rowCount} filas insertadas exitosamente.`);
                resolve({ success: true, insertedRows: rowCount });
            }
        }
    );

    // Definir los tipos de datos para un rendimiento y seguridad óptimos.
    request.addParameter('item', TYPES.Int);
    request.addParameter('n_partida', TYPES.NVarChar);
    request.addParameter('nombre_del_articulo', TYPES.NVarChar);
    request.addParameter('articulo_descripcion', TYPES.NVarChar);
    request.addParameter('colores', TYPES.NVarChar);
    request.addParameter('cod_colores', TYPES.Int);
    request.addParameter('pqt', TYPES.Int);
    request.addParameter('kg', TYPES.Decimal, { precision: 18, scale: 2 }); // Ideal para números con decimales.

    // Tedious espera un array de arrays para el bulk load.
    data.forEach(item => {
        request.addRow(
            item.item,
            item.n_partida,
            item.nombre_del_articulo,
            item.articulo_descripcion,
            item.colores,
            item.cod_colores,
            item.pqt,
            item.kg
        );
    });

    try {
        connection.execSql(request);
    } catch (error) {
        console.error("Error al ejecutar el SQL del Bulk Insert:", error);
        resolve({ success: false, error: `Error al ejecutar la inserción: ${error.message}` });
        connection.close();
    }
}


/**
 * Manejador NUEVO para obtener información del sistema.
 */
ipcMain.handle('get-system-info', () => {
    try {
        return {
            success: true,
            data: {
                cpuCount: os.cpus().length,
                totalMemory: (os.totalmem() / (1024 ** 3)).toFixed(2) + ' GB', // Memoria en GB
            }
        };
    } catch (error) {
        return { success: false, error: "No se pudo obtener la información del sistema." };
    }
});

// ---> FIN: NUEVOS MANEJADORES Y LÓGICA DE BASE DE DATOS <---

// ======================================================================
// === FIN: MANEJADORES DE IPC ==========================================
// ======================================================================



// ======================================================================
// === INICIO: CÓDIGO DE ARRANQUE DE LA APP (SIN CAMBIOS) ================
// ======================================================================

app.setName('Procesador de Excel');
const gotTheLock = app.requestSingleInstanceLock();

if (!gotTheLock) {
    app.quit();
} else {
    app.on('second-instance', () => {
        const windows = BrowserWindow.getAllWindows();
        if (windows.length) {
            if (windows[0].isMinimized()) windows[0].restore();
            windows[0].focus();
        }
    });

    app.whenReady().then(() => {
        createWindow();
        app.on('window-all-closed', () => {
            if (process.platform !== 'darwin') app.quit();
        });
        app.on('activate', () => {
            if (BrowserWindow.getAllWindows().length === 0) createWindow();
        });
    });

    process.on('uncaughtException', (error) => {
        console.error('Error no capturado:', error);
        dialog.showErrorBox('Error', `Se produjo un error inesperado: ${error.message}`);
    });

    process.on('unhandledRejection', (reason) => {
        console.error('Promesa rechazada no manejada:', reason);
        dialog.showErrorBox('Error', `Error en la aplicación: ${reason.message || reason}`);
    });
}
// ======================================================================
// === FIN: CÓDIGO DE ARRANQUE DE LA APP (SIN CAMBIOS) ===================
// ======================================================================