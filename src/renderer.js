// I removed the qualifications and exclusions textboxes for the sake of keeping things as simple as possible to start

let save_ready = false;
let csv_ready = false;

let con = document.getElementById("console");

let startBTN = document.getElementById('start-btn');

let saveBTN = document.getElementById('save-location-btn');
let saveLabel = document.getElementById('save-location-label');

let contractorinput = document.getElementById("contractor-input");
let jobinput = document.getElementById("job-input");
let dateinput = document.getElementById("date-input");
// let qualificationsinput = document.getElementById("qualifications-input");
// let exclusionsinput = document.getElementById("exclusions-input");
// let priceinput = document.getElementById("price-input");

let contractorvalue = "";
let jobvalue = "";
let datevalue = "";
// let qualificationsvalue = "";
// let exclusionsvalue = "";
// let pricevalue = "";

contractorinput.addEventListener("change", (e) => { contractorvalue = e.target.value });
jobinput.addEventListener("change", (e) => { jobvalue = e.target.value });
dateinput.addEventListener("change", (e) => { datevalue = e.target.value });
// priceinput.addEventListener("change", (e) => { pricevalue = e.target.value });

document.getElementById("save-location-btn").addEventListener("click", (e) => {
    document.getElementById('save-loader').classList.remove('hide');
    document.getElementById('save-loader').classList.add('loading');

    window.electronAPI.saveLocation();
})

let csvBTN = document.getElementById("open-file-btn");

csvBTN.addEventListener("click", (e) => {
    document.getElementById('csv-loader').classList.remove('hide');
    document.getElementById('csv-loader').classList.add('loading');
    window.electronAPI.openFile();
})

window.electronAPI.saveLocationSuccess((success, location) => {
    document.getElementById('save-loader').classList.add('hide');
    document.getElementById('save-loader').classList.remove('loading');
    if(success) {
        saveBTN.classList.add("success");
        document.getElementById('save-success-marker').innerText = '✔️'
        saveLabel.innerHTML = `
            <span class="save-label-label">saving proposal to: </span><span class="save-label-location">${location}</span>
        `
        save_ready = true;
        checkStart();
    } else {
        // tell user the save location is invalid
    }
})

let csvLabel = document.getElementById("generation-success-label");

window.electronAPI.csvLocationSuccess((success, location) => {
    document.getElementById('csv-loader').classList.add('hide');
    document.getElementById('csv-loader').classList.remove('loading');
    if(success) {
        csvBTN.classList.add("success");
        document.getElementById('csv-success-marker').innerText = '✔️'
        csvLabel.innerHTML = `
            <span class="save-label-label">using csv located at: </span><span class="save-label-location">${location}</span>
        `
        csv_ready = true;
        checkStart();
    } else {
        // tell user the save location is invalid
    }
})

let successWrapper = document.getElementById('success-wrapper');

window.electronAPI.proposalSuccess((success) => {
    if(success) {
        successWrapper.hidden = false;
        setTimeout(() => {
            successWrapper.hidden = true;
        }, 1500)
    } else {

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
                <div class="settings-label main-name">name:<span class="tooltip-name">your name - used in the signature section of the proposal</span></div><input id="settings-name" class="settings-input" type="text" placeholder="John Smith" value="${currentSettings.info.user}">
                <div class="settings-label main-primary-column">primary column:<span class="tooltip-primary-column">what column has your unique tool info? ex: 'base cabinets w/ sub tops'</span></div>
                <select id="settings-primary-column" class="settings-input">
                    <option value="Subject" ${currentSettings.settings.primary_column == "Subject" ? 'Selected' : ''}>Subject</option>
                    <option value="Label" ${currentSettings.settings.primary_column == "Label" ? 'Selected' : ''}>Label</option>
                </select>
                <div class="settings-label main-generate-details">use legacy details:<span class="tooltip-generate-details">use the old style totals sheet?</span></div><input id="settings-generate-details" class="settings-input" type="checkbox" ${currentSettings.settings.generate_details ? 'checked' : ''}>
                <div class="settings-label main-open-word">auto open word:<span class="tooltip-open-word">open proposals automatically in word after creation?</span></div><input id="settings-open-word" class="settings-input" type="checkbox" ${currentSettings.settings.auto_open_word ? 'checked' : ''}>
                <div class="settings-label main-top-shop">topshop tool subject:<span class="tooltip-top-shop">what subject, label, or secondary column name is being used for top shop items?       ! use a comma to filter multiple subjects ex: Quartz tops, Solid Surface tops</span></div><input id="settings-top-shop-subject" class="settings-inputs" type="text" placeholder="countertops" value="${currentSettings.settings.top_shop_tool_subject}">
                <div class="settings-label disabled-label">require all fields:</div><input id="settings-require-inputs" class="settings-inputs disabled-input" type="checkbox" disabled ${currentSettings.settings.require_all_fields ? 'checked' : ''}>
                <button id="settings-save-btn" class="settings-save-btn">save</button>
            `

            let saveName = document.getElementById('settings-name');
            let primaryColumn = document.getElementById('settings-primary-column');
            let generateDetails = document.getElementById('settings-generate-details');
            let saveOpenWord = document.getElementById('settings-open-word');
            let saveTopShopToolSubject = document.getElementById('settings-top-shop-subject');
            let saveRequireInputs = document.getElementById('settings-require-inputs');

            saveName.addEventListener('change', (e) => { currentSettings.info.user = e.target.value });
            primaryColumn.addEventListener('change', (e) => { currentSettings.settings.primary_column = e.target.selectedIndex == 1? "Label" : "Subject" });
            generateDetails.addEventListener('change', (e) => { console.log(e); currentSettings.settings.generate_details = e.target.checked });
            saveOpenWord.addEventListener('change', (e) => { console.log(e); currentSettings.settings.auto_open_word = e.target.checked });
            saveTopShopToolSubject.addEventListener('change', (e) => { currentSettings.settings.top_shop_tool_subject = e.target.value });
            saveRequireInputs.addEventListener('change', (e) => { currentSettings.settings.require_all_inputs = e.target.checked });

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

let checkStart = () => {
    if(save_ready && csv_ready) {
        startBTN.classList.remove('btn-disabled');
    }
}

startBTN.onclick = (e) => {
    if(contractorvalue.length > 0 && jobvalue.length > 0 && datevalue.length > 0 && save_ready && csv_ready) {
        window.electronAPI.start({
            contractor: contractorvalue,
            job: jobvalue,
            date: datevalue,
        });
        // document.getElementById("open-file-btn").classList.add("success");
        document.getElementById("errors").innerText = ""
    } else {
        // document.getElementById("errors").innerText = "! please enter a contractor, job name, and bid date to continue"
    }
}

// to finish settings just send the settings object back to main and save it

window.electronAPI.getSettings();