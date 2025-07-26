const { contextBridge, ipcRenderer } = require('electron');

// Expone un API segura y bien definida al proceso de renderizado.
// Cada función aquí es un "puente" que invoca una operación en el proceso principal.
contextBridge.exposeInMainWorld('api', {
  /**
   * Pide al proceso principal que abra un diálogo de selección de archivo.
   * @returns {Promise<string|null>} La ruta del archivo o null si se cancela.
   */
  openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),

  /**
   * Pide al proceso principal que procese un archivo Excel y lo exporte.
   * @param {string} filePath - La ruta del archivo a procesar.
   * @returns {Promise<Object>} El resultado de la operación.
   */
  processAndExport: (filePath) => ipcRenderer.invoke('process-excel-and-export', filePath),

  /**
   * Pide al proceso principal que lea un archivo Excel y devuelva los datos limpios.
   * (Esta función es para el paso 2: Enviar a BD)
   * @param {string} filePath - La ruta del archivo a leer.
   * @returns {Promise<Array>} Un array de objetos con los datos de las filas.
   */
  readCleanedExcel: (filePath) => ipcRenderer.invoke('read-cleaned-excel', filePath),

  /**
   * Pide al proceso principal que envíe los datos a la base de datos.
   * @param {Array} data - Los datos a insertar.
   * @returns {Promise<Object>} El resultado de la inserción.
   */
  sendDataToDB: (data) => ipcRenderer.invoke('send-data-to-db', data),

  /**
   * Obtiene información básica del sistema desde el proceso principal.
   * @returns {Promise<Object>} Un objeto con información como cpus y memoria.
   */
  getSystemInfo: () => ipcRenderer.invoke('get-system-info')
});