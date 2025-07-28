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
// qualificationsinput.addEventListener("change", (e) => { qualificationsvalue = e.target.value });
// exclusionsinput.addEventListener("change", (e) => { exclusionsvalue = e.target.value });
priceinput.addEventListener("change", (e) => { pricevalue = e.target.value });

document.getElementById("save-location-btn").addEventListener("click", (e) => {
    window.electronAPI.saveLocation();
})

document.getElementById("open-file-btn").addEventListener("click", (e) => {
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