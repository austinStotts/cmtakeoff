if(require('electron-squirrel-startup')) return;
const { app, BrowserWindow, ipcMain, dialog } = require('electron/main');
const path = require('node:path');
let csvMethods = require("./index.js");
let Settings = require("./settings.js");

// you will see similar code block all over
// this just tests to see if the program is being run from a dev environment
// if the program running the code is 'vscode' then we are not in production...
// things like opening work or the default save location are changed to be quicker for testing 
// also most of the file paths change when installed vs run from this script
let devMode = false;
if(process.env.TERM_PROGRAM == 'vscode') {
  devMode = true;
}

// holder variables so all the methods have access to these
let saveLocation = "";
let csvLocation = "";
let mainWindow;

// this file is the heart of the program
// despite not being all that long it creates the window and renders the html
// as well as handling all communication to the client
// the middle man if you will


// does what it says on the tin
const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 400,
    height: 600,
    icon: process.resourcesPath + "../images/logo.png",
    
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })
  mainWindow.loadFile('src/index.html');
  !devMode ? mainWindow.menuBarVisible = false : mainWindow.menuBarVisible = true;
}

// runs when a user clicks start
let handleStart = (event, details) => {
  let callback = (success) => {
    mainWindow.webContents.send('proposal-success', true);
  };
  csvMethods.calculateProposal(csvLocation, details, saveLocation, settings, handleError, callback);
}

// runs when a user clicks select csv
let openFile = (event) => {
  console.log("opening file");
  let filePath = dialog.showOpenDialog({
    title: 'Select CSV you created in blubeam',
    buttonLabel: 'Open',
    filters: [{ name: "CSV", extensions: ['csv'] }],
    properties: ['openFile']
  }).then(result =>  {
    console.log(result.canceled);
    csvLocation = result.filePaths[0];
    mainWindow.webContents.send('csv-location-success', true, csvLocation);
  }).catch(err => {
    console.log(err);
    mainWindow.webContents.send('csv-location-success', false, csvLocation);
  })
}

// runs when a user clicks select save loaction
let handleSaveLocation = (event) => {
  dialog.showSaveDialog({
    title: 'Select location to save proposal',
    defaultPath: devMode ? "C:\\Users\\astotts\\Desktop\\CSV TESTING\\proposal.docx" : "P:\\Plans Download\\proposal.docx",
    buttonLabel: 'Save',
    filters: [{ name: "DOCX", extensions: ['docx'] }],
    properties: []
  }).then(file => {
    console.log(file.canceled);
    if (!file.canceled) {
      console.log('got the file!');
      console.log(file.filePath.toString());
      saveLocation = file.filePath.toString();
      mainWindow.webContents.send('save-location-success', true, saveLocation);
    }
  }).catch(err => {
    console.log(err);
    mainWindow.webContents.send('save-location-success', false);
  });
}

// sends an error code to the client to be handled there
let handleError = (error, code) => {
  mainWindow.webContents.send('error', error, code);
}

// runs when the client wants the current saved settings
let handleGetSettings = (event) => {
  mainWindow.webContents.send('return-settings', true, settings);
}

// runs when the client sends a new settings object
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

// runs on startup to define the api endpoints
app.whenReady().then(() => {
  createWindow();
  ipcMain.on('open-file', openFile);
  ipcMain.on('save-location', handleSaveLocation);
  ipcMain.on('get-settings', handleGetSettings);
  ipcMain.on('set-settings', handleSetSettings);
  ipcMain.on('start', handleStart);
})

// runs iff all the windows are closed to end the program
// remove to make one of those annoying apps that you have to force quit to actually close
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

// the proposal generation is not working right...