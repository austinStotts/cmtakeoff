const { contextBridge, ipcRenderer, ipcMain } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: (details) => ipcRenderer.send('open-file', details),
  saveLocation: () => ipcRenderer.send('save-location'),
  saveLocationSuccess: (callback) => { ipcRenderer.on('save-location-success', (event, success, location) => callback(success, location)) },
  csvLocationSuccess: (callback) => { ipcRenderer.on('csv-location-success', (event, success, location) => callback(success, location)) }
})
