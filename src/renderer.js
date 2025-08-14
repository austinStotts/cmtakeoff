// I removed the qualifications and exclusions textboxes for the sake of keeping things as simple as possible to start

let con = document.getElementById("console");



let saveBTN = document.getElementById('save-location-btn');
let saveLabel = document.getElementById('save-location-label');

let contractorinput = document.getElementById("contractor-input");
let jobinput = document.getElementById("job-input");
let dateinput = document.getElementById("date-input");
// let qualificationsinput = document.getElementById("qualifications-input");
// let exclusionsinput = document.getElementById("exclusions-input");
let priceinput = document.getElementById("price-input");

let contractorvalue = "";
let jobvalue = "";
let datevalue = "";
// let qualificationsvalue = "";
// let exclusionsvalue = "";
let pricevalue = "";

contractorinput.addEventListener("change", (e) => { contractorvalue = e.target.value });
jobinput.addEventListener("change", (e) => { jobvalue = e.target.value });
dateinput.addEventListener("change", (e) => { datevalue = e.target.value });
priceinput.addEventListener("change", (e) => { pricevalue = e.target.value });

document.getElementById("save-location-btn").addEventListener("click", (e) => {
    window.electronAPI.saveLocation();
})

let csvBTN = document.getElementById("open-file-btn");

csvBTN.addEventListener("click", (e) => {
    if(contractorvalue.length > 0 && jobvalue.length > 0 && datevalue.length > 0 && pricevalue.length > 0) {
        window.electronAPI.openFile({
            contractor: contractorvalue,
            job: jobvalue,
            date: datevalue,
            // qualifications: qualificationsvalue,
            // exclusions: exclusionsvalue,
            price: pricevalue,
        });
        // document.getElementById("open-file-btn").classList.add("success");
        document.getElementById("errors").innerText = ""
    } else {
        document.getElementById("errors").innerText = "! please enter a contractor, job name, and bid date to continue"
    }
})

window.electronAPI.saveLocationSuccess((success, location) => {
    if(success) {
        saveBTN.classList.add("success");
        saveLabel.innerHTML = `
            <span class="save-label-label">saving proposal to: </span><span class="save-label-location">${location}</span>
        `
    } else {
        // tell user the save location is invalid
    }
})

let csvLabel = document.getElementById("generation-success-label");

window.electronAPI.csvLocationSuccess((success, location) => {
    if(success) {
        csvBTN.classList.add("success");
        csvLabel.innerHTML = `
            <span class="save-label-label">successfully saved proposal to: </span><span class="save-label-location">${location}</span>
        `
    } else {
        // tell user the save location is invalid
    }
})

window.electronAPI.error((error, code) => {
    window.alert(error);
    switch(code) {
        case('0001'):
            console.log('');
    }
})


let currentSettings;

window.electronAPI.returnSettings((success, data) => {
    if(success) {
        currentSettings = data;
    } else {
        // could not get settings from main
    }
})

let view = document.getElementById('main-view-wrapper');
let settingsView = document.getElementById('settings-view-wrapper');

let currentView = 'home'
let settingsBtn = document.getElementById("settings-btn").addEventListener('click', (e) => {
    console.log('clicking settings');
    // send request to main for current settings
    window.electronAPI.getSettings();
    switch(currentView) {
        case('home'):
            currentView = 'settings';
            view.hidden = true;
            settingsView.innerHTML = `
                <div class="settings-label">name:</div><input id="settings-name" class="settings-input" type="text" placeholder="John Smith" value="${currentSettings.info.user}">
                <div class="settings-label disabled-label">auto open word:</div><input id="settings-open-word" class="settings-input disabled-input" type="checkbox" disabled ${currentSettings.settings.auto_open_word ? 'checked' : ''}>
                <div class="settings-label disabled-label">require all fields:</div><input id="settings-require-inputs" class="settings-inputs disabled-input" type="checkbox" disabled ${currentSettings.settings.require_all_fields ? 'checked' : ''}>
                <div class="settings-label disabled-label">remember inputs:</div><input id="settings-remember-inputs" class="settings-inputs disabled-input" type="checkbox" disabled ${currentSettings.settings.remember_inputs ? 'checked' : ''}>
                <button id="settings-save-btn" class="settings-save-btn">save</button>
            `

            let saveName = document.getElementById('settings-name');
            let saveOpenWord = document.getElementById('settings-open-word');
            let saveRequireInputs = document.getElementById('settings-require-inputs');
            let saveRememberInputs = document.getElementById('settings-remember-inputs');

            saveName.addEventListener('change', (e) => { currentSettings.info.user = e.target.value });
            saveOpenWord.addEventListener('change', (e) => { console.log(e); currentSettings.settings.auto_open_word = e.target.checked });
            saveRequireInputs.addEventListener('change', (e) => { currentSettings.settings.require_all_inputs = e.target.checked });
            saveRememberInputs.addEventListener('change', (e) => { currentSettings.settings.remember_inputs = e.target.checked });

            let saveSettingsBTN = document.getElementById('settings-save-btn');
            saveSettingsBTN.addEventListener('click', (e) => {
                window.electronAPI.setSettings(currentSettings);
            })
            window.electronAPI.setSettingsSuccess((success) => {
                console.log('inside success renderer' , success)
                if(success) {
                    saveSettingsBTN.innerText = 'SAVED';
                    saveSettingsBTN.classList.add('disabled-btn');
                    setTimeout(() => {
                        saveSettingsBTN.innerText = 'save';
                        saveSettingsBTN.classList.remove('disabled-btn');
                    }, 2000)
                } else {
                    saveSettingsBTN.innerText = 'ERROR';
                    console.log('error in renderer success', success)
                }
            })
            break
        case('settings'):
            currentView = 'home';
            view.hidden = false;
            settingsView.innerHTML = '';
            break
    }
})

// to finish settings just send the settings object back to main and save it

window.electronAPI.getSettings();