const fs = require('fs')

let settingsLocation = process.resourcesPath + `/saves/settings.json`;
let devMode = false;
if(process.env.TERM_PROGRAM == 'vscode') {
  devMode = true;
  settingsLocation = './saves/settings.json'
}

let loadSettings = (callback) => {
    return fs.readFile(settingsLocation, 'utf8', (error, data) => {
        if(error) {
            if(error.errno == -4058) {
                console.log("settings does not exist... creating settings.json");
                fs.writeFile(settingsLocation, JSON.stringify(defaults), (error, data) => {
                    if(error) {
                        console.log(error);
                        return false;
                    } else {
                        console.log('successfully saved defualt settings');
                        callback(defaults);
                    }
                })
            } else {
                console.log(error);
            }
        } else {
            let result = JSON.parse(data);
            callback(result);
        }
    })
}

let updateSetting = (settings) => 
    new Promise((resolve, reject) => {
        fs.writeFile(settingsLocation, JSON.stringify(settings), (err) => {
            if (err) reject(err, false)
            else resolve(true)
        })
    })








let defaults = {
    "info": {
        "user": "guest"
    },
    "settings": {
        "primary_column": "Subject",
        "generate_details": false, // this is called 'use legacy details' on the client ...also... is not currently user changable 
        "auto_open_word": true,
        "add_page_labels": true,
        "include_default_exclusions": true,
        "show_notes": true,
        "top_shop_tool_subject": "tops",
        "require_all_fields": true,
    }
}


module.exports = { loadSettings, updateSetting }