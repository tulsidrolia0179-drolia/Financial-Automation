const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

const customMappingPath = path.join(app.getPath('userData'), 'custom_mapping.json');
if (!fs.existsSync(customMappingPath)) {
    fs.writeFileSync(customMappingPath, '{}');
}

function createWindow () {
    const mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
        }
    });

    // We will load a local HTML file for the UI shell
    mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
    ipcMain.handle('dialog:selectFile', async () => {
        const { canceled, filePaths } = await dialog.showOpenDialog({
            properties: ['openFile'],
            filters: [
                { name: 'Documents', extensions: ['xlsx', 'xls', 'pdf'] }
            ]
        });
        if (!canceled && filePaths.length > 0) {
            return filePaths[0];
        }
        return null;
    });

    createWindow();

    app.on('activate', function () {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});
