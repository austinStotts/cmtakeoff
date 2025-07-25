const { app, BrowserWindow, ipcMain, dialog } = require("electron/main");
let fs = require("fs");
const PDFDocument = require("pdfkit");
const { parse } = require("csv-parse");

let roundall = (scope) => {
  let keys = Object.keys(scope);
  for (let i = 0; i < keys.length; i++) {
    let key = keys[i];
    let items = Object.keys(scope[key]);
    for (let j = 0; j < items.length; j++) {
      scope[key][items[j]][0] = Math.ceil(scope[key][items[j]][0]);
    }
  }
};

let rows = [];

let calculateProposal = (file, details, saveLocation) => {
  if (file != null) {
    fs.createReadStream(file)
      .pipe(parse({ delimiter: ",", from_line: 2 }))
      .on("data", function (row) {
        rows.push(row);
      })
      .on("end", () => {
        console.log("input successful");
        let scope = {};
        let totals = {};
        // This loop organizes the raw csv data into rooms and items
        // it is mandatory for the csv columns to be in the correct order
        // scope - measurement - label - subject - special comments
        for (let i = 0; i < rows.length; i++) {
          let room = rows[i][0];
          let measurement = rows[i][1];
          let unit = rows[i][2];
          let item = rows[i][3];
          let label = rows[i][4];
          // scope
          if (!scope[room]) {
            scope[room] = {};
            scope[room][item] = [Number(measurement), unit];
          } else {
            if (scope[room][item] == undefined) {
              scope[room][item] = [Number(measurement), unit];
            } else {
              scope[room][item][0] =
                Number(scope[room][item][0]) + Number(measurement);
            }
          }
          // totals
          console.log(room);
          let roomarray = room.split(" ");
          let n = 1;
          if (roomarray[0].toLowerCase().includes("x)")) {
            n = roomarray[0].slice(1, roomarray[0].length - 2);
          } else if (roomarray[0].toLowerCase().includes("(x")) {
            n = roomarray[0].slice(2, roomarray[0].length - 1);
          } else if (
            roomarray[roomarray.length - 1].toLowerCase().includes("x)")
          ) {
            n = roomarray[roomarray.length - 1].slice(
              1,
              roomarray[roomarray.length - 1].length - 2
            );
          } else if (
            roomarray[roomarray.length - 1].toLowerCase().includes("(x")
          ) {
            n = roomarray[roomarray.length - 1].slice(
              2,
              roomarray[roomarray.length - 1].length - 1
            );
          } else {
          }

          if (totals[item]) {
            totals[item] = totals[item] + Number(measurement) * n;
          } else {
            totals[item] = Number(measurement) * n;
          }
        }

        // before the scope is used in the document all numbers are rounded up
        roundall(scope);
        console.log(totals);

        // PDF Creation
        // adds the header date and other data from the boxes in the application
        const doc = new PDFDocument({ size: "LETTER",});
        doc.pipe(fs.createWriteStream(saveLocation));
        doc.fontSize(24);
        doc.font("Times-Bold").text("Custom Millwork", { align: "center", columns: 1 });
        doc.fontSize(16);
        doc.font("Times-Bold");
        doc.text("1200 Murphy Drive", { align: "center" });
        doc.text("Maumelle, AR 72113", { align: "center" });
        doc.text("Phone: 501.851.4421", { align: "center" });
        let imageWidth = 70;
        doc.image('./AWICQW.png',
          doc.page.width/2 - imageWidth/2,doc.y,{
          width:imageWidth,
        });
        doc.moveDown(4);
        doc.fontSize(12);
        doc.text(`Bid Date: ${details.date}`);
        doc.moveDown(1);
        doc.text(`To: ${details.contractor}`);
        doc.text(`Project: ${details.job}`);
        doc.text(`Attention: Estimating`);
        // doc.moveDown(2);
        doc.text("\n");
        doc.text("\n");

        // qualifications
        doc.text("Pricing Qualifications:", { underline: true });
        doc.font("Times-Roman");
        //         doc.text(`- Drawer pulls:   4-inch wire pull
        // Door Pulls:   4-inch wire pull
        // Hinges:       Blum concealed self-close
        // Drawers:      Blum Metabox system`, {  })
        details.qualifications.split("\n").forEach((qualification) => {
          doc.text(qualification);
        });
        // doc.moveDown(2);
        doc.text("\n");
        doc.text("\n");

        // exclusions
        doc.font("Times-Bold");
        doc.text("Project Specific Exclusions:", { underline: true });
        doc.font("Times-Roman");
        //         doc.text(`Excludes QCP Certification and Labels
        // Excludes Arkansas Wage Rate/LEED/FSC certification labels
        // Excludes All Lighting and Electrical
        // Excludes All Painting and Finishing
        // Excludes All Demolition
        // `, {  })
        doc.font("Times-Roman");
        details.exclusions.split("\n").forEach((exclusion) => {
          doc.text(exclusion);
        });
        // doc.moveDown(2);
        doc.text("\n");
        doc.text("\n");

        // Main Room and Item loop
        // the outer loop (i) is looping through each room
        // the inner loop (j) is looping through the items in that room
        let i = 1;
        for (const [key, value] of Object.entries(scope)) {
          // (i) Rooms
          doc.font("Times-Bold");
          doc.text(`${i}. ${key}`);
          i++;
          for (let [room, item] of Object.entries(value)) {
            // (j) Items
            doc.font("Times-Roman");
            doc.fontSize(12);
            // depending on the 'unit' of measurement the appropriate description will be used
            if (room.startsWith("LF of")) {
              room = room.slice(5);
            } // attempt to filter out the use of proposal text in the tools
            if (item[1] == "sf") {
              doc.text(`${item[0]} SQFT of ${room}`, { indent: "8" });
            } else if (item[1] == "Count") {
              doc.text(`(${item[0]}X) ${room}`, { indent: "8" });
            } else {
              doc.text(`${item[0]} LF of ${room}`, { indent: "8" });
            }
          }
        }

        // bid ammount line
        // the collection of symbols in the replace function are a regular expresion to format the number with commas
        // doc.moveDown(2);
        doc.text("\n");
        doc.text("\n");
        doc.text(
          `Base Price for Furnished and Installed: ......$${details.price
            .toString()
            .replace(/\B(?=(\d{3})+(?!\d))/g, ",")}.00`,
          {}
        );
        // doc.moveDown(2);
        doc.text("\n");
        doc.text("\n");
        doc.text(`
NOTE:\n
Installation of any item (s) not supplied by Custom Millwork are Excluded\n
EXCLUSIONS: Anything not listed above including but not limited to the following:\n
\n
1.  Finishing
2.  Sinks (Unless Otherwise Noted)
3.  Wall, Floor & Ceiling Tile
4.  All after Hours Work
5.  Standing or Running Trims Unless Otherwise Listed
6.  All Acoustic Wall &/or Fabric Covered Panels
7.  All Tackable &/or Writing Boards
8.  All Blocking Unless Otherwise Noted
9.  Glass and Glass Hardware Unless Otherwise Noted
10. Hollow Metal Work
11. Finish or Rough Hardware
12. Interior & Exterior Trim Unless Otherwise Noted
13. Metal or Wood Framing
14. Performance and Payment Bonds
15. Liquidated Damages
16. Steel or Wood doors, jambs and Casings
17. Overtime work due to job falling behind schedule ahead of our scope.
`);

        // doc.moveDown(2);
        doc.text("\n");
        doc.text("\n");
        doc.text(`
CONDITIONS:
1.  Upon acceptance, as evidenced by signatures of the purchaser and an officer of our company, this proposal becomes a valid contract.  If other contract forms are used, this proposal automatically becomes a part of any contract into which we might enter for the work covered – whether or not the contract stated that wood work will be a defined in this quotation.
2.  Proposal is good for 14 days.
3.  Work in progress cannot be changed without a written change order.
4.  No back charges will be allowed without our written consent.
5.  Clerical errors subject to correction.
6.  Required delivery dates shall be given in writing.  Should job be unable to take delivery at scheduled time, we shall be entitled to payment for such materials upon presentation of proper certification and insurance.
`);

        doc.end();

        // save a second document of the same name + '-totals'
        // show totals for each item in the job

        let saveLocation2 = saveLocation.split(".")[0] + "-totals.pdf";
        const doc2 = new PDFDocument({ size: "LETTER" });
        doc2.pipe(fs.createWriteStream(saveLocation2));
        doc2.fontSize(16);
        doc2.font("Times-Roman");
        for (let [item, value] of Object.entries(totals)) {
          doc2.text(`${item}: ${Math.ceil(value)}`);
        }
        doc2.end();
      });
  } else {
    console.log("cannot parse csv - no file location");
  }
};

module.exports = { calculateProposal };
