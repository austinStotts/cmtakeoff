if(require('electron-squirrel-startup')) return;
const { app, BrowserWindow, ipcMain, dialog } = require('electron/main');
const path = require('node:path');
let csvMethods = require("./index.js");


let mainWindow;

const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 400,
    height: 600,
    icon: process.resourcesPath + "/images/logo.png",
    
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })
  mainWindow.loadFile('src/index.html')
}


let saveLocation = "";

let openFile = (event, details) => {
  console.log("opening file")
  dialog.showOpenDialog({properties: ['openFile']})
  .then(result =>  {
    console.log(result.canceled)
    csvMethods.calculateProposal(result.filePaths[0], details, saveLocation);
  }).catch(err => {
    console.log(err)
  })
}

let handleSaveLocation = (event) => {
  dialog.showSaveDialog({
    title: 'Select the File Path to save',
    defaultPath: path.join(__dirname, './proposal.pdf'),
    buttonLabel: 'Save',
    filters: [{extensions: ['pdf']}],
    properties: []
  }).then(file => {
    console.log(file.canceled);
    if (!file.canceled) {
      console.log('got the file!')
      console.log(file.filePath.toString());
      saveLocation = file.filePath.toString();
      mainWindow.webContents.send('save-location-success', true, saveLocation);
    }
  }).catch(err => {
      console.log(err);
      mainWindow.webContents.send('save-location-success', false);
  });
}

app.whenReady().then(() => {
  createWindow();
  ipcMain.on('open-file', openFile);
  ipcMain.on('save-location', handleSaveLocation);
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})


