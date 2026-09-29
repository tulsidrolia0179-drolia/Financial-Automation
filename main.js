const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { extractPreviousYearData } = require('./processor.js');

function createWindow() {
    const win = new BrowserWindow({
        width: 1000,
        height: 600,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false
        }
    });
    win.loadFile('index.html');
}

app.whenReady().then(() => {
    createWindow();
    ipcMain.handle('dialog:selectFile', async () => {
        const result = await dialog.showOpenDialog({
            properties: ['openFile'],
            filters: [{ name: 'Documents', extensions: ['xlsx', 'xls', 'pdf'] }]
        });
        if (result.canceled) return null;
        return result.filePaths[0];
    });
    ipcMain.handle('process:file', async (event, filePath, taskType) => {
        if (taskType === '3') {
            return await extractPreviousYearData(filePath);
        }
        return { success: true, message: `File recognized for module ${taskType}` };
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});
