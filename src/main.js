if(require('electron-squirrel-startup')) return;
const { app, BrowserWindow, ipcMain, dialog } = require('electron/main');
const path = require('node:path');
let csvMethods = require("./index.js");

let devMode = false;
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
    let callback = () => {
      mainWindow.webContents.send('csv-location-success', true, saveLocation);
    }
    csvMethods.calculateProposal(result.filePaths[0], details, saveLocation, callback);
  }).catch(err => {
    console.log(err)
  })
}

let handleSaveLocation = (event) => {
  dialog.showSaveDialog({
    title: 'Select the File Path to save',
    defaultPath: devMode ? "C:\\Users\\astotts\\Desktop\\CSV TESTING\\proposal.docx" : "P:\\Plans Download\\proposal.docx",
    buttonLabel: 'Save',
    filters: [{name: "DOCX", extensions: ['docx']}],
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


// maybe add auto open in word