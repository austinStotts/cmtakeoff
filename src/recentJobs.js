const fs = require('fs')

let recentsLocation = process.resourcesPath + `/saves/settings.json`;
let devMode = false;
if(process.env.TERM_PROGRAM == 'vscode') {
  devMode = true;
  recentsLocation = './saves/recents.json'
}

let getRecentJobs = () => {
    fs.readFile(recentsLocation, (error, data) => {
        if(error) {

        } else {
            
        }
    })
}