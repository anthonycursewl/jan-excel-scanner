const isApiReady = () => {
    return window.api && window.api.openFileDialog && window.api.processExcelAndExport;
};

const colors = {
    info: 'blue',
    success: 'green',
    warning: 'orange',
    error: 'red',
    default: 'black'
};

const createStatusMessage = () => {
    const statusMessage = document.createElement('p');
    statusMessage.id = 'statusMessage';
    document.body.appendChild(statusMessage);
    return statusMessage;
};

const showStatus = (message, type = 'info') => {
    const statusMessage = document.getElementById('statusMessage') || createStatusMessage();
    statusMessage.textContent = message;
    statusMessage.style.color = colors[type] || colors.default;
    statusMessage.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
};

const handleTabularClick = async () => {
    console.log("Iniciando proceso de tabulación...");
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

        if (result?.success) {
            showStatus(`¡Proceso completado! Archivo limpio guardado en: ${result.outputPath}`, 'success');
            console.log('Proceso finalizado. Archivo limpio guardado en:', result.outputPath);
            return result;
        } else {
            throw new Error(result?.error || 'Error desconocido durante el procesamiento');
        }
    } catch (error) {
        const errorMessage = error.message || 'Error desconocido';
        showStatus(`Error: ${errorMessage}`, 'error');
        console.error('Error al procesar el archivo:', error);
        throw error;
    }
};

const initTabular = () => {
    document.addEventListener('DOMContentLoaded', () => {
        const tabularButton = document.getElementById('Tabular');

        if (!tabularButton) {
            console.error('Error: No se encontró el botón con id "Tabular".');
            return;
        }

        console.log('Botón "Tabular" encontrado. Configurando manejador de eventos.');

        tabularButton.addEventListener('click', async () => {
            if (!isApiReady()) {
                const errorMsg = 'Error: La API de Electron no está disponible. ¿Está cargado el preload correctamente?';
                console.error(errorMsg);
                showStatus(errorMsg, 'error');
                return;
            }

            try {
                await handleTabularClick();
            } catch (error) {
                console.error('Error en el proceso de tabulación:', error);
            }
        });
    });
};

initTabular();