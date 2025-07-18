const { app, BrowserWindow, ipcMain, dialog } = require('electron/main');
let fs = require("fs");
const PDFDocument = require('pdfkit');
const { parse } = require("csv-parse");

let roundall = (scope) => {
  let keys =  Object.keys(scope);
  for(let i = 0; i < keys.length; i++) {
    let key = keys[i]
    let items = Object.keys(scope[key]);
    for(let j = 0; j < items.length; j++) {
      scope[key][items[j]] = Math.ceil(scope[key][items[j]])
    }
  }
}

let organize = (scope) => {
  let keys = Object.keys(scope);
  keys.forEach(key => {
    scope[key].name = key
  })
}

// let csv = "../../../../Desktop/Marshall County Revised.csv";
let rows = [];

let calculateProposal = (file, details, saveLocation) => {
  if(file != null) {
    fs.createReadStream(file)
      .pipe(parse({ delimiter: ",", from_line: 2 }))
      .on("data", function (row) {
        rows.push(row);
    }).on("end", () => {
        console.log("input successful");
        let scope = {}
        // console.log(rows)
        for(let i = 0; i < rows.length; i++) {
          console.log(rows[i])
          if(!scope[rows[i][0]]) {
            scope[rows[i][0]] = {}
            scope[rows[i][0]][rows[i][3]] = Number(rows[i][1])
          } else {
            if(scope[rows[i][0]][rows[i][3]] == undefined) {
              scope[rows[i][0]][rows[i][3]] = Number(rows[i][1]);
            } else {
              scope[rows[i][0]][rows[i][3]] = Number(scope[rows[i][0]][rows[i][3]]) + Number(rows[i][1]);  
            }
          }

        }

        // console.log(scope)
        roundall(scope);
        // organize(scope);
        console.log(scope);
        


        // PDF CREATE
        const doc = new PDFDocument({ size: 'LETTER' });
        doc.pipe(fs.createWriteStream(saveLocation)); // write to PDF
        // doc.pipe(res);                                       // HTTP response
        // doc.addPage({ size: "LETTER"})
        doc.fontSize(24)
        doc.font("Times-Bold").text("Custom Millwork", { align: 'center' })
        doc.fontSize(16)
        doc.font("Times-Bold").text("1200 Murphy Drive", { align: 'center' })
        doc.font("Times-Bold").text("Maumelle, AR 72113", { align: 'center' })
        doc.font("Times-Bold").text("Phone: 501.851.4421", { align: 'center' })
        // doc.image("./AWICQW.png", { width: '100', fit: [200, 200], align: 'center',  })
        let imageWidth = 100 // what you wants
        doc.image('./AWICQW.png', 
          doc.page.width/2 - imageWidth/2,doc.y,{
          width:imageWidth
        });
        doc.moveDown(8)
        doc.fontSize(12)
        doc.text(`Bid Date: ${details.date}`, { align: 'left' })
        doc.moveDown(1)
        doc.text(`To:               ${details.contractor}`, { align: 'left' })
        doc.text(`Project:       ${details.job}`, { align: 'left' })
        doc.text(`Attention:   Estimating`, { align: 'left' })
        doc.moveDown(2);
        doc.text('Pricing Qualifications:', { underline: true });
        details.qualifications.split('\n').forEach(qualification => {
          doc.text(qualification);
        })
        doc.moveDown(2);
        doc.text('Project Specific Exclusions:', { underline: true });
        details.exclusions.split('\n').forEach(exclusion => {
          doc.text(exclusion);
        })
        doc.moveDown(2);
        let i = 1;
        for (const [key, value] of Object.entries(scope)) {
          doc.font('Times-Bold');
          doc.text(`${i}: ${key}`, )
          i++
          for (const [room, item] of Object.entries(value)) {
            doc.font('Times-Roman');
            doc.text(`${item} LF of ${room}`, { indent: '36' })
          }

        }

        doc.moveDown(2);
        doc.text(`Base Price for Furnished and Installed: ......$${details.price.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",")}.00`)
        doc.moveDown(2);
        doc.text(`
NOTE:
Installation of any item (s) not supplied by Custom Millwork are Excluded
EXCLUSIONS: Anything not listed above including but not limited to the following: 

1.   Finishing
2.   Sinks (Unless Otherwise Noted)
3.   Wall, Floor & Ceiling Tile
4.   All after Hours Work
5.   Standing or Running Trims Unless Otherwise Listed
6.   All Acoustic Wall &/or Fabric Covered Panels 
7.   All Tackable &/or Writing Boards
8.   All Blocking Unless Otherwise Noted
9.   Glass and Glass Hardware Unless Otherwise Noted
10.  Hollow Metal Work
11.  Finish or Rough Hardware 
12.  Interior & Exterior Trim Unless Otherwise Noted
13.  Metal or Wood Framing
14.  Performance and Payment Bonds
15.  Liquidated Damages
16.  Steel or Wood doors, jambs and Casings
17.  Overtime work due to job falling behind schedule ahead of our scope.
`)



        doc.end();



    })
  } else {
    console.log("cannot parse csv - no file location")
  }

}


module.exports = { calculateProposal }