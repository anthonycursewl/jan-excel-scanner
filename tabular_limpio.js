function isApiReady() {
    return window.api && window.api.openFileDialog && window.api.processExcelAndExport;
}

function showStatus(message, type = 'info') {
    const colors = {
        info: 'blue',
        success: 'green',
        warning: 'orange',
        error: 'red',
        default: 'black'
    };
    
    const statusMessage = document.getElementById('statusMessage') || createStatusMessage();
    statusMessage.textContent = message;
    statusMessage.style.color = colors[type] || colors.default;
    
    // Hacer scroll al mensaje
    statusMessage.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// Función para crear el elemento de estado si no existe
function createStatusMessage() {
    const statusMessage = document.createElement('p');
    statusMessage.id = 'statusMessage';
    document.body.appendChild(statusMessage);
    return statusMessage;
}

// Función para manejar el clic en el botón
async function handleTabularClick() {
    console.log("buenas tarddfñlkdlñsfd")
    showStatus('Iniciando selección y procesamiento de Excel...', 'info');

    try {
        const filePath = await window.api.openFileDialog();

        if (!filePath) {
            showStatus('Selección de archivo cancelada.', 'warning');
            console.warn('Selección de archivo cancelada.');
            return;
        }

        showStatus(`Archivo seleccionado: ${filePath}. Procesando...`, 'info');
        console.log('Archivo seleccionado:', filePath);

        const result = await window.api.processExcelAndExport(filePath);

        if (result && result.success) {
            showStatus(`¡Proceso completado! Archivo limpio guardado en: ${result.outputPath}`, 'success');
            console.log('Proceso finalizado. Archivo limpio guardado en:', result.outputPath);
        } else {
            const errorMessage = result?.error || 'Error desconocido durante el procesamiento';
            showStatus(`Error durante el procesamiento: ${errorMessage}`, 'error');
            console.error('Error durante el procesamiento del archivo:', errorMessage);
        }
    } catch (error) {
        const errorMessage = error.message || 'Error desconocido';
        showStatus(`Error crítico: ${errorMessage}`, 'error');
        console.error('Error al abrir diálogo o procesar archivo:', error);
    }
}

// Inicialización cuando el DOM esté listo
// tabular_limpio.js (Versión simplificada para depuración)
document.addEventListener('DOMContentLoaded', () => {
    const tabularButton = document.getElementById('Tabular');
    const statusMessage = document.getElementById('statusMessage');

    if (tabularButton) {
        console.log('Botón "Tabular" encontrado. Añadiendo listener.');

        tabularButton.addEventListener('click', async () => {
            console.log('¡Botón clickeado!'); // <-- ¿Ves esto en la consola?

            // Comprobamos si la API existe JUSTO antes de usarla
            if (typeof window.api?.openFileDialog !== 'function') {
                const errorMsg = 'Error Crítico: window.api.openFileDialog no es una función. ¿El script preload.js cargó correctamente?';
                console.error(errorMsg);
                statusMessage.textContent = errorMsg;
                statusMessage.style.color = 'red';
                return;
            }

            statusMessage.textContent = 'Abriendo selector de archivos...';
            statusMessage.style.color = 'blue';

            try {
                const filePath = await window.api.openFileDialog();

                if (filePath) {
                    statusMessage.textContent = `Procesando: ${filePath}`;
                    const result = await window.api.processAndExport(filePath);
                    if (result.success) {
                        statusMessage.textContent = `¡Éxito! Guardado en: ${result.outputPath}`;
                        statusMessage.style.color = 'green';
                    } else {
                        throw new Error(result.error);
                    }
                } else {
                    statusMessage.textContent = 'Operación cancelada.';
                    statusMessage.style.color = 'orange';
                }
            } catch (error) {
                console.error('Error en el proceso:', error);
                statusMessage.textContent = `Error: ${error.message}`;
                statusMessage.style.color = 'red';
            }
        });

    } else {
        console.error('Error: No se encontró el botón con id "Tabular".');
    }
});