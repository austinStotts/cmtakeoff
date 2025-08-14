const { contextBridge, ipcRenderer, ipcMain } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: (details) => ipcRenderer.send('open-file', details),
  saveLocation: () => ipcRenderer.send('save-location'),
  getSettings: () => ipcRenderer.send('get-settings'),
  setSettings: (settings) => ipcRenderer.send('set-settings', settings),
  saveLocationSuccess: (callback) => { ipcRenderer.on('save-location-success', (event, success, location) => callback(success, location)) },
  csvLocationSuccess: (callback) => { ipcRenderer.on('csv-location-success', (event, success, location) => callback(success, location)) },
  returnSettings: (callback) => { ipcRenderer.on('return-settings', (event, success, data) => callback(success, data)) },
  setSettingsSuccess: (callback) => { ipcRenderer.on('set-settings-success', (event, success) => callback(success)) },
  error: (callback) => { ipcRenderer.on('error', (event, error) => callback(error)) },
})
