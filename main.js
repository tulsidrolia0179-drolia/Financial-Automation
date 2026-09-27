const { app, BrowserWindow } = require('electron');
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
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  // We will load a local HTML file for the UI shell
  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});
