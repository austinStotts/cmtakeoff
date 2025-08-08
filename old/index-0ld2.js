const { app, BrowserWindow, ipcMain, dialog } = require("electron/main");
let fs = require("fs");
const PDFDocument = require("pdfkit");
const { parse } = require("csv-parse");

let devMode = true;

let roundall = (scope) => {
  let keys = Object.keys(scope);
  for (let i = 0; i < keys.length; i++) {
    let key = keys[i];
    let items = Object.keys(scope[key]);
    for (let j = 0; j < items.length; j++) {
      if(items[j] != "pages") { // do not round the page array
        scope[key][items[j]][0] = Math.round(scope[key][items[j]][0] + 0.2);
      }
    }
  }
};

let randomsort = (a, b) => 0.5 - Math.random();

//                   array       string
let checkPages = (currentPages, page) => {
  let hasPage = false;
  for(let i = 0; i < currentPages.length; i++) {
    console.log(currentPages[i], page);
    if(currentPages[i] == page) { hasPage = true }
  }
  return hasPage;
}

// [
//   'Space',            'Measurement',
//   'Measurement Unit', 'Subject',
//   'Page Label',       'Length',
//   'Length Unit',      'Depth',
//   'Depth Unit',       'Label',
//   'Room Multi',       'Top Footage',
//   'Total Footages',   'Material Type'
// ]

let rows = [];

let calculateProposal = (file, details, saveLocation, callback) => {
  console.log("file: ", file);
  console.log("details: ", details);
  console.log("saveLocation: ", saveLocation);
  if (file != null) {
    fs.createReadStream(file)
      .pipe(parse({ delimiter: ",", from_line: 1 }))
      .on("data", function (row) {
        rows.push(row);
      })
      .on("end", () => {
        console.log("input successful");
        let columns = rows.shift();
        columns[0] = columns[0].slice(1);
        console.log(columns);
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
          let page = rows[i][4];
          let label = rows[i][5];
          let row = {};
          for(let j = 0; j < columns.length; j++) {
            row[columns[j]] = rows[i][j];
          }
          console.log(row);

          // console.log(room, measurement, unit, item, label);
          // scope
          if (!scope[room]) { // if the room does not exist
            if (page) { scope[room] = { pages: [page] } }
            scope[room][item] = [Number(measurement), unit];
            // scope[room].pages.push(`${page}`);
          } else { // if the item does not exist
            if (scope[room][item] == undefined) {
              scope[room][item] = [Number(measurement), unit];
              if(page && !checkPages(scope[room].pages, page)) {
                scope[room].pages.push(`${page}`);
              }
            } else { // if the item already exists
              scope[room][item][0] = Number(scope[room][item][0]) + Number(measurement);
              if(page && !checkPages(scope[room].pages, page)) {
                scope[room].pages.push(`${page}`);
              }
            }
          }
          // totals
          // console.log(room);
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
            totals[item] = [totals[item][0] + Number(measurement) * n, unit];
          } else {
            totals[item] = [Number(measurement) * n, unit];
          }
        }

        // before the scope is used in the document all numbers are rounded up
        roundall(scope);
        console.log(scope);

        // PDF Creation
        // adds the header date and other data from the boxes in the application
        const doc = new PDFDocument({ size: "LETTER", margin: 72});
        doc.pipe(fs.createWriteStream(saveLocation));
        doc.text(`                                                                                                                                            `)
        doc.fontSize(24);
        doc.font("Times-Bold").text("Custom Millwork", { align: "center"});
        doc.fontSize(16);
        doc.font("Times-Bold");
        // doc.text("1200 Murphy Drive", { align: "center" });
        // doc.text("Maumelle, AR 72113", { align: "center" });
        // doc.text("Phone: 501.851.4421", { align: "center" });
        let headerText = ["1200 Murphy Drive", "Maumelle, AR 72113", "Phone: 501.851.4421"];
        doc.text(`${headerText.map(item => item).join("\n")}`, { align: 'center' })
        let imageWidth = 70;

        if(devMode) {
          doc.image("./images/AWICQW.png",
            doc.page.width/2 - imageWidth/2,doc.y,{
            width:imageWidth,
          });
        } else {
          doc.image(process.resourcesPath + '/images/AWICQW.png',
            doc.page.width/2 - imageWidth/2,doc.y,{
            width:imageWidth,
          });
        }




        doc.text(`



  `);

        doc.fontSize(12);
        doc.text(details.date);
        doc.text(`
          `)
        let jobText = [`To: ${details.contractor}`, `Project: ${details.job}`, `Attention: Estimating
          `];
        doc.text(jobText.join("\n"));
        // doc.text(`
        //   `);
      
        doc.font("Times-Bold");
        doc.text(`Pricing Qualifications:`, { underline: true });
        let qualificationText = [`Drawer pulls: 4-inch wire pull`, `Door Pulls: 4-inch wire pull`, `Hinges: Blum concealed self-close`, `Drawers: Blum Metabox system
          `];
        doc.font("Times-Roman");
        doc.text(qualificationText.join("\n"));
        // doc.text(`
        //   `);

        
        doc.font("Times-Bold");
        doc.text(`Project Exclusions:`, { underline: true });
        let exclusionText = [`Excludes QCP Certification and Labels`, `Excludes Arkansas Wage Rate/LEED/FSC certification labels`, `Excludes All Lighting and Electrical`, `Excludes All Painting and Finishing`, `Excludes All Demolition
          `];
        doc.font("Times-Roman");
        doc.text(exclusionText.join("\n"));
        // doc.text(`
          
        //   `);
        
        
        let i = 1;
        for (const [key, value] of Object.entries(scope)) {
          // (i) Rooms
          doc.font("Times-Bold");
          doc.text(`${i}. ${key}`);
          i++;
          let itemlist = [];
          for (let [room, item] of Object.entries(value)) {
            // (j) Items

            // depending on the 'unit' of measurement the appropriate description will be used
            if (room.startsWith("LF of")) {
              room = room.slice(5);
            } else if (room.startsWith("of")) {
              room = room.slice(2);
            } // attempt to filter out the use of proposal text in the tools
            if (room == "pages") {
              doc.text(`As seen on: (${item.join(', ')})`);
            } else if (item[1] == "sf") {
              itemlist.push(`${item[0]} SQFT of ${room}`);
            } else if (item[1] == "Count") {
              itemlist.push(`(${item[0]}X) ${room}`);
            } else {
              itemlist.push(`${item[0]} LF of ${room}`);
            }
          }
          doc.font("Times-Roman");
          doc.fontSize(12);
          doc.text(`${itemlist.join('\n')}
          `)
        }

        doc.text(`
  `)
        doc.font("Times-Bold");
        doc.text(
          `Base Price for Furnished and Installed: ......$${details.price
            .toString()
            .replace(/\B(?=(\d{3})+(?!\d))/g, ",")}.00`,
          {}
        );

        doc.text(`
          
          `)

        let noteText1 = [
          `NOTE:`,
          `Installation of any item(s) not supplied by Custom Millwork are Excluded`,
          `EXCLUSIONS: Anything not listed above including but not limited to the following:`,
        ]
        let noteText2 = [
          `1.  Finishing`,
          `2.  Sinks (Unless Otherwise Noted)`,
          `3.  Wall, Floor & Ceiling Tile`,
          `4.  All after Hours Work`,
          `5.  Standing or Running Trims Unless Otherwise Listed`,
          `6.  All Acoustic Wall &/or Fabric Covered Panels`,
          `7.  All Tackable &/or Writing Boards`,
          `8.  All Blocking Unless Otherwise Noted`,
          `9.  Glass and Glass Hardware Unless Otherwise Noted`,
          `10. Hollow Metal Work`,
          `11. Finish or Rough Hardware`,
          `12. Interior & Exterior Trim Unless Otherwise Noted`,
          `13. Metal or Wood Framing`,
          `14. Performance and Payment Bonds`,
          `15. Liquidated Damages`,
          `16. Steel or Wood doors, jambs and Casings`,
          `17. Overtime work due to job falling behind schedule ahead of our scope
          `,
        ];


        doc.font("Times-Bold");
        doc.fontSize(12);
        doc.text(noteText1.join("\n"));
        doc.font("Times-Roman");
        doc.fontSize(10);
        doc.text(noteText2.join("\n"));
        doc.text(`
          
          `)

        let conditionsText = [
          `1. Upon acceptance, as evidenced by signatures of the purchaser and an officer of our company, this proposal becomes a valid contract.  If other contract forms are used, this proposal automatically becomes a part of any contract into which we might enter for the work covered – whether or not the contract stated that wood work will be a defined in this quotation.`,
          `2. Proposal is good for 14 days.`,
          `3. Work in progress cannot be changed without a written change order.`,
          `4. No back charges will be allowed without our written consent.`,
          `5. Clerical errors subject to correction.`,
          `6. Required delivery dates shall be given in writing.  Should job be unable to take delivery at scheduled time, we shall be entitled to payment for such materials upon presentation of proper certification and insurance.
          `
        ]
        
        doc.font("Times-Bold");
        doc.fontSize(12);
        doc.text(`CONDITIONS:`);
        doc.font("Times-Roman");
        doc.fontSize(10);
        doc.text(conditionsText.join("\n"));
        doc.text(`
          
          `)

        
        let importantText1 = [
          `1. Cabinets, brand named or kitchen`,
          `2. Chalk, cork or bulletin boards (except wood trim, furnished loose, unmitered)`,
          `3. Compositions, except caps for columns & items to be shop fastened to millwork items`,
          `4. Fabrics, felt or soft plastics`,
          `5. Fencing materials`,
          `6. Flooring, deck boards & catwalks`,
          `7. Folding or sliding brand name doors`,
          `8. Frames and picket doors for folding & overhead doors`,
          `9. Glass, glazing or mirrors`,
          `10.  Hard boards`,
          `11.  Door hardware, not the preparation for installation of`,
          `12.  Insulation, cement or asbestos materials`,
          `13.  Labor at job site`,
          `14.  Ladders `,
          `15.  Linoleum`,
          `16.  Metal items of any kind`,
          `17.  Metal louvers, grilles, or their installation`,
          `18.  Metal moldings used with materials supplied by others`,
          `19.  Overhead doors or their jambs`,
          `20.  Plastic items other than laminated plastic`,
          `21.  Priming, back-painting finishing or preservative treatment`,
          `22.  Repairs or alterations to existing materials`,
          `23.  Roof decking, exposed`,
          `24.  Shop notching or cutting on shelving & battens furnished loose`,
          `25.  Temporary millwork`,
          `26.  Weather-stripping`,
          `27.  Wood shingles, shakes or siding`,
          `28.  Wood, rough bucks, furring, & other materials not exposed to view`,
          `29.  X-ray & refrigerator doors`,
          `30.  Dock bumper`,
          `31.  Cash allowances
          `,
        ]

        let importantText2 = [
          `(B)  Flush plywood panel work furnished in commercial size panels unless shop work is required on edges.`,
          `(C)  Terms are net, due the fifteenth of the month following requisition of materials. We reserve the right to stop delivery of materials if payments are not made when due. Payments not made in full by the by the 26th proximal will subject the balance to a finance charge of 1.5% per calendar month or fraction thereof until paid.`,
          `(D)  All contracts are contingent upon strikes, breakdowns, fires, or other causes beyond our control and are accepted subject to approved credit.
          `,
        ]

        doc.font("Times-Bold");
        doc.fontSize(12);
        doc.text(`(A.)  IMPORTANT: THE FOLLOWING ITEMS ARE CONSIDERED 
“ARCHITECTURAL WOODWORK” AND DO NOT FORM A PART OF THE 
PROPOSAL UNLESS NOTED:`)
        doc.font("Times-Roman");
        doc.fontSize(10);
        doc.text(importantText1.join("\n"));
        doc.font("Times-Bold");
        doc.fontSize(12);
        doc.text(importantText2.join("\n"))

        doc.end();

        // save a second document of the same name + '-totals'
        // show totals for each item in the job

        let saveLocation2 = saveLocation.split(".")[0] + "-totals.pdf";
        const doc2 = new PDFDocument({ size: "LETTER", margins: { top: 25, bottom: 25, left: 50, right: 50 } });
        doc2.pipe(fs.createWriteStream(saveLocation2));
        doc2.fontSize(10);
        doc2.font("Courier-Bold");
        doc2.text(details.job);
        doc2.text(saveLocation2);
        doc2.font("Courier");
        let data = [];
        // console.log(totals);
        for (let [item, value] of Object.entries(totals)) {
          let row = [item, value[1] == `ft' in"` ? "LF" : value[1] == `ft` ? "LF" : value[1], Math.ceil(value[0])];
          data.push(row);
        }
        data.sort();
        doc2.table({
          data,
          columnStyles: ["*", 50, 50],
        })
        doc2.end();
        callback()
      });
  } else {
    console.log("cannot parse csv - no file location");
  }
};

module.exports = { calculateProposal };
