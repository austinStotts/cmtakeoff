if(require('electron-squirrel-startup')) return;
const { app, BrowserWindow, ipcMain, dialog } = require('electron/main');
const path = require('node:path');
let csvMethods = require("./index.js");
let Settings = require("./settings.js");

let devMode = false;
if(process.env.TERM_PROGRAM == 'vscode') {
  devMode = true;
}
let mainWindow;
// var settings;

const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 400,
    height: 600,
    icon: process.resourcesPath + "/images/logo.png",
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })
  mainWindow.loadFile('src/index.html');
}

let saveLocation = "";

let openFile = (event, details) => {
  console.log("opening file")
  dialog.showOpenDialog({properties: ['openFile']})
  .then(result =>  {
    console.log(result.canceled)
    let callback = (success) => {
      mainWindow.webContents.send('csv-location-success', success, saveLocation);
    }
    console.log('setting in main at line 36', settings);
    csvMethods.calculateProposal(result.filePaths[0], details, saveLocation, settings, handleError, callback);
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

let handleError = (error, code) => {
  mainWindow.webContents.send('error', error, code);
}

let handleGetSettings = (event) => {
  mainWindow.webContents.send('return-settings', true, settings);
}

let handleSetSettings = async (event, newSettings) => {
  Settings.updateSetting(newSettings)
  .then(success => {
    console.log(success);
    mainWindow.webContents.send('set-settings-success', true);
    Settings.loadSettings(saveSettings);
  }).catch((error, success) => {
    console.log(error, success);
    mainWindow.webContents.send('set-settings-success', false);
  })
}

app.whenReady().then(() => {
  createWindow();
  ipcMain.on('open-file', openFile);
  ipcMain.on('save-location', handleSaveLocation);
  ipcMain.on('get-settings', handleGetSettings);
  ipcMain.on('set-settings', handleSetSettings);
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

let settings;

let saveSettings = (data) => {
  settings = data
}

Settings.loadSettings(saveSettings);




// maybe add auto open in word - done ✔️

// finish settings
// add settings screen
// add more error handling
// recent jobs
// batch jobs
// make n copies of proposals with different details
// walter does not get a totals sheet when generating ✔️
// remove sqft from proposal and just use measurement ✔️

// later - rework the client with react and allow for more screens
// setup screen that allows for removing room/items from proposal/totals


// test for a bunch of different things
// test if room labels being on or off works
// test for depths being on and off
// test for anything else relating to the changes made since last friday