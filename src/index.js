const { app, BrowserWindow, ipcMain, dialog } = require("electron/main");
let fs = require("fs");
const PDFDocument = require("pdfkit");
const { parse } = require("csv-parse");
const open = require('open');
const docx = require("docx");

let generateDoc = true; // true
let generateTotals = true; // true
let isSuccess = true;

let devMode = false;
if(process.env.TERM_PROGRAM == 'vscode') {
  devMode = true;
}


// [
//   'Space',            'Measurement',
//   'Measurement Unit', 'Subject',
//   'Page Label',       'Length',
//   'Length Unit',      'Depth',
//   'Depth Unit',       'Label',
//   'Room Multi',       'Top Footage',
//   'Total Footages',   'Material',
//   'Group',
// ]



let rows = [];

let calculateProposal = (file, details, saveLocation, settings, error, callback) => {
  console.log("file: ", file);
  console.log("details: ", details);
  console.log("saveLocation: ", saveLocation);
  let stats;
  fs.stat(file, (error, data) => {
    if(error) {
      console.log(error);
    } else {
      stats = data;
    }
  })
  if (file != null) {
    fs.createReadStream(file)
      .pipe(parse({ delimiter: ",", from_line: 1 }))
      .on("data", function (row) {
        rows.push(row);
      })
      .on("end", () => {
        isSuccess = true;
        console.log("input successful");
        console.log(settings);
        let columns = rows.shift();
        columns[0] = columns[0].slice(1);
        console.log(columns);
        let toolKey = settings.settings.primary_column || "Subject";
        let searchKey = toolKey == "Subject" ? "Label" : "Subject";
        let totals = { groups: {} };
        let job = { groups: {} };
        for (let i = 0; i < rows.length; i++) {
          let row = {};
          for(let j = 0; j < columns.length; j++) {
            row[columns[j]] = rows[i][j];
          }

          if(!row['Group']) { row['Group'] = 'Base Bid' } // a blank group will be changed to 'Base Bid'

          // Job
          // loops through each row of the csv and organizes into job > groups > rooms > items
          // be careful changing refs to the job object because many things rely on it
          // pass by ref vs pass by val
          if(job.groups[row['Group']]) {
            if(job.groups[row['Group']].rooms[row['Space']]) {
              if(job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`]) { // old item
                job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`].measurement += Number(row['Measurement']);
              } else { // new item
                job.groups[row['Group']].rooms[row['Space']].pages.push(row['Page Label']);
                job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`] = {
                  name: row[toolKey],
                  search: row[searchKey],
                  measurement: Number(row['Measurement']),
                  unit: row['Measurement Unit'],
                  material: row['Material'] || '*',
                  depth: Number(row['Depth']) || null,
                  groupName: row['Group'],
                  spaceName: row['Space'],
                  excluded: false
                }
              }
            } else { // new space
              job.groups[row['Group']].rooms[row['Space']] = { items: {}, pages: [] };
              job.groups[row['Group']].rooms[row['Space']].pages.push(row['Page Label']);
              job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`] = {
                name: row[toolKey],
                search: row[searchKey],
                measurement: Number(row['Measurement']),
                unit: row['Measurement Unit'],
                material: row['Material'] || '*',
                depth: Number(row['Depth']) || null,
                groupName: row['Group'],
                spaceName: row['Space'],
                excluded: false
              }
            }
          } else { // new group
            job.groups[row['Group']] = { rooms: {} };
            job.groups[row['Group']].rooms[row['Space']] = { items: {}, pages: [] };
            job.groups[row['Group']].rooms[row['Space']].pages.push(row['Page Label']);
            job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`] = {
              name: row[toolKey],
              search: row[searchKey],
              measurement: Number(row['Measurement']),
              unit: row['Measurement Unit'],
              material: row['Material'] || '*',
              depth: Number(row['Depth']) || null,
              groupName: row['Group'],
              spaceName: row['Space'],
              excluded: false
            }
          }
        }


        // TOTALS
        // uses the job object to make a totals object
        // when saving refs to the job a deepcopy is used (JSON.parse(JSON.stringify(x)))
        // this lets the document maker make changes in place to the job object without changing the totals
        Object.keys(job.groups).forEach(group => {
          if(!totals.groups[group]) {
            totals.groups[group] = { items: {} }
          }

          Object.keys(job.groups[group].rooms).forEach(room => {
            let roomarray = room.split(" ");
            let n = 1;
            if (roomarray[0].toLowerCase().includes("x)")) {
              n = roomarray[0].slice(1, roomarray[0].length - 2);
            } else if (roomarray[0].toLowerCase().includes("(x")) {
              n = roomarray[0].slice(2, roomarray[0].length - 1);
            } else if (roomarray[roomarray.length - 1].toLowerCase().includes("x)")) {
              n = roomarray[roomarray.length - 1].slice(1, roomarray[roomarray.length - 1].length - 2);
            } else if (roomarray[roomarray.length - 1].toLowerCase().includes("(x")) {
              n = roomarray[roomarray.length - 1].slice(2, roomarray[roomarray.length - 1].length - 1);
            } else {
              // none
            }

            Object.keys(job.groups[group].rooms[room].items).forEach(item => {
              if(totals.groups[group].items[item]) { // old item
                if(job.groups[group].rooms[room].items[item].depth) {
                  totals.groups[group].items[item].measurement += (job.groups[group].rooms[room].items[item].measurement * (job.groups[group].rooms[room].items[item].depth / 12)) * n;
                  totals.groups[group].items[item].details.push([job.groups[group].rooms[room].items[item].spaceName, (job.groups[group].rooms[room].items[item].measurement * (job.groups[group].rooms[room].items[item].depth / 12)) * n, 'sqft', JSON.parse(JSON.stringify(job.groups[group].rooms[room].items[item])), job.groups[group].rooms[room].pages]);
                } else {
                  totals.groups[group].items[item].measurement += job.groups[group].rooms[room].items[item].measurement * n;
                  totals.groups[group].items[item].details.push([job.groups[group].rooms[room].items[item].spaceName, job.groups[group].rooms[room].items[item].measurement * n, job.groups[group].rooms[room].items[item].unit, JSON.parse(JSON.stringify(job.groups[group].rooms[room].items[item])), job.groups[group].rooms[room].pages]);
                }
              } else { // new item
                if(job.groups[group].rooms[room].items[item].depth) { // depth should be in inches
                  totals.groups[group].items[item] = { measurement: (job.groups[group].rooms[room].items[item].measurement * (job.groups[group].rooms[room].items[item].depth / 12)) * n, unit: 'sqft', details: [[job.groups[group].rooms[room].items[item].spaceName, (job.groups[group].rooms[room].items[item].measurement * (job.groups[group].rooms[room].items[item].depth / 12)) * n, 'sqft', JSON.parse(JSON.stringify(job.groups[group].rooms[room].items[item])), job.groups[group].rooms[room].pages]] }
                } else {
                  totals.groups[group].items[item] = { measurement: job.groups[group].rooms[room].items[item].measurement * n, unit: job.groups[group].rooms[room].items[item].unit, details: [[job.groups[group].rooms[room].items[item].spaceName, job.groups[group].rooms[room].items[item].measurement * n, job.groups[group].rooms[room].items[item].unit, JSON.parse(JSON.stringify(job.groups[group].rooms[room].items[item])), job.groups[group].rooms[room].pages]] }
                }
              }
            })
          }) 
        })


        
        // Word DOCX
        // uses the job object to generate text for the proposal
        // lines needed in the proposal are pushing into the scopelist array
        // each index in the scopelist array is a new line on the proposal
        let awiimg = '';
        let iconimg = '';
        if(devMode) {
          awiimg = "./images/AWICQW.png";
          iconimg = "./images/icon.jpg";
        } else {
          awiimg = process.resourcesPath + '/images/AWICQW.png'
          iconimg = process.resourcesPath + '/images/icon.jpg'
        }

        let scopelist = [];

        Object.keys(job.groups).forEach(group => {
          scopelist.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: ``,
                size: 24,
              }),
            ]
          }))
          scopelist.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `${group}`,
                bold: true,
                size: 24,
                underline: true,
              })
            ]
          }))
          let rooms = Object.keys(job.groups[group].rooms);
          for(let room = 0; room < rooms.length; room++) { // rooms
            let items = Object.keys(job.groups[group].rooms[rooms[room]].items);
            scopelist.push(new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: `${rooms[room]} ${job.groups[group].rooms[rooms[room]].pages.length > 0 ? 'as seen on (' + [...new Set(job.groups[group].rooms[rooms[room]].pages)].join(", ") + ')' : ''}`,
                  bold: true,
                  size: 24,
                })
              ],
              numbering: {
                reference: "my-numbering",
                level: 0,
              }
            }));

            // tops should read as LF on the proposal while all other items with depth should be sqft
            for(let item = 0; item < items.length; item++) { // items
              let thing = job.groups[group].rooms[rooms[room]].items[items[item]]
              if(thing.depth > 0) { // this double if statement is for the special tops case
                let topshopSubjects = settings.settings.top_shop_tool_subject.split(",");
                topshopSubjects.push("tops", "countertops");
                let match = false;
                topshopSubjects.forEach(subject => {
                  subject = subject.trim();
                  console.log(subject);
                  if (thing.search == subject) match = true;
                })
                if(match) {} // keep as lf
                else { // changes all non top items to sqft !WARNING !THIS CHANGES THE JOB OBJECT AND ALL REFS TO THE JOB OBJECT!
                  job.groups[group].rooms[rooms[room]].items[items[item]].measurement = (job.groups[group].rooms[rooms[room]].items[items[item]].measurement * (job.groups[group].rooms[rooms[room]].items[items[item]].depth / 12));
                  job.groups[group].rooms[rooms[room]].items[items[item]].unit = 'sqft'
                  console.log(job.groups[group].rooms[rooms[room]].items[items[item]])
                }
              }
              if(job.groups[group].rooms[rooms[room]].items[items[item]].unit == 'Count') {
                scopelist.push(new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `(${Math.ceil(job.groups[group].rooms[rooms[room]].items[items[item]].measurement)}X) ${items[item]}`,
                      size: 24,
                    })
                  ],
                  numbering: {
                    reference: "my-bullets",
                    level: 0,
                  }
                }));
              } else if (job.groups[group].rooms[rooms[room]].items[items[item]].unit == 'sqft' || job.groups[group].rooms[rooms[room]].items[items[item]].unit == 'sf') {
                scopelist.push(new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `${Math.ceil(job.groups[group].rooms[rooms[room]].items[items[item]].measurement)} SQFT of ${items[item]}`,
                      size: 24,
                    })
                  ],
                  numbering: {
                    reference: "my-bullets",
                    level: 0,
                  }
                }));
              } else {
                scopelist.push(new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `${Math.ceil(job.groups[group].rooms[rooms[room]].items[items[item]].measurement)} LF of ${items[item]}`,
                      size: 24,
                    })
                  ],
                  numbering: {
                    reference: "my-bullets",
                    level: 0,
                  }
                }));
              }
            }
          }
          scopelist.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: ``,
                size: 24,
              }),
            ]
          }))
          scopelist.push(new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: `Price to Furnish and Install ${group} from Above:......$`,
                  bold: true,
                  size: 24,
                })
              ],
            })
          )
        })

        // it is easier to have these large blocks of text broken out and formatted into arrays
        // each index of the array will be its own paragraph

        let notelist = []
        let noteText1 = [
          `NOTE:`,
          `Installation of any item(s) not supplied by Custom Millwork are Excluded`,
          `EXCLUSIONS: Anything not listed above including but not limited to the following:`,
        ]
        let noteText2 = [
          `1. Finishing`,
          `2. Sinks (Unless Otherwise Noted)`,
          `3. Wall, Floor & Ceiling Tile`,
          `4. All after Hours Work`,
          `5. Standing or Running Trims Unless Otherwise Listed`,
          `6. All Acoustic Wall &/or Fabric Covered Panels`,
          `7. All Tackable &/or Writing Boards`,
          `8. All Blocking Unless Otherwise Noted`,
          `9. Glass and Glass Hardware Unless Otherwise Noted`,
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

        noteText1.forEach(note => {
          notelist.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: note,
                size: 24,
                bold: true,
              })
            ]
          }))
        })

        noteText2.forEach(note => {
          notelist.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: note,
                size: 20,
              })
            ]
          }))
        })

        let conditionslist = [];
        let conditionsText = [
          `1. Upon acceptance, as evidenced by signatures of the purchaser and an officer of our company, this proposal becomes a valid contract.  If other contract forms are used, this proposal automatically becomes a part of any contract into which we might enter for the work covered – whether or not the contract stated that woodwork will be a defined in this quotation.`,
          `2. Proposal is good for 14 days.`,
          `3. Work in progress cannot be changed without a written change order.`,
          `4. No back charges will be allowed without our written consent.`,
          `5. Clerical errors subject to correction.`,
          `6. Required delivery dates shall be given in writing.  Should job be unable to take delivery at the scheduled time, we shall be entitled to payment for such materials upon presentation of proper certification and insurance.
          `
        ]

        conditionsText.forEach(condition => {
          conditionslist.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: condition,
                size: 20,
              })
            ]
          }))
        })


        let importantlist1 = [];
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
          `10. Hard boards`,
          `11. Door hardware, not the preparation for installation of`,
          `12. Insulation, cement or asbestos materials`,
          `13. Labor at job site`,
          `14. Ladders `,
          `15. Linoleum`,
          `16. Metal items of any kind`,
          `17. Metal louvers, grilles, or their installation`,
          `18. Metal moldings used with materials supplied by others`,
          `19. Overhead doors or their jambs`,
          `20. Plastic items other than laminated plastic`,
          `21. Priming, back-painting finishing or preservative treatment`,
          `22. Repairs or alterations to existing materials`,
          `23. Roof decking, exposed`,
          `24. Shop notching or cutting on shelving & battens furnished loose`,
          `25. Temporary millwork`,
          `26. Weather-stripping`,
          `27. Wood shingles, shakes or siding`,
          `28. Wood, rough bucks, furring, & other materials not exposed to view`,
          `29. X-ray & refrigerator doors`,
          `30. Dock bumper`,
          `31. Cash allowances
          `,
        ]

        let importantlist2 = []
        let importantText2 = [
          `(B.)  Flush plywood panel work is furnished in commercial size panels unless shop work is required on edges.`,
          `(C.)  Terms are net, due the fifteenth of the month following requisition of materials. We reserve the right to stop delivery of materials if payments are not made when due. Payments not made in full by the by the 26th proximal will subject the balance to a finance charge of 1.5% per calendar month or fraction thereof until paid.`,
          `(D.)  All contracts are contingent upon strikes, breakdowns, fires, or other causes beyond our control and are accepted subject to approved credit.
          `,
        ]


        
        importantText1.forEach(text => {
          importantlist1.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: text,
                size: 20,
              })
            ]
          }))
        })
        
        importantText2.forEach(text => {
          importantlist2.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: text,
                size: 24,
                bold: true,
              })
            ], 
            spacing: {
              after: 30
            }
          }))
        })
        











        // Main Proposal
        // the body of the proposal
        // all the paragraphs are needed to format the document correctly
        // paragraphs in word are the main way to get a true 'new line'
        // text and textruns with new lines in them are not the same as a real paragraph break
        // the arrays of text prepared above are added in place

        // this bit is using the 'docx' package
        // all text sizes are 2x the size you would pick in word
        // size 24 => 12pt font

        // I used extra spaces to approximate an equal distant look but it isnt perfect
        const doc = new docx.Document({
          numbering: {
            config: [
              {
                reference: "my-numbering",
                levels: [
                  {
                    level: 0,
                    format: docx.LevelFormat.DECIMAL,
                    alignment: docx.AlignmentType.LEFT,
                    text: '%1.',
                    style: {
                      paragraph: {
                        indent: { left: docx.convertInchesToTwip(0.25), hanging: docx.convertInchesToTwip(0.25),  },
                      },
                        
                    },
                  },
                ],
              },
                  {
                    reference: "my-bullets",
                    levels: [
                        {
                            level: 0,
                            format: docx.LevelFormat.BULLET,
                            alignment: docx.AlignmentType.LEFT,
                            text: '•',
                            style: {
                                paragraph: {
                                    indent: { left: docx.convertInchesToTwip(0.25), hanging: docx.convertInchesToTwip(0.25),  },
                                },
                            
                            },
                        },
                    ],
                },
              ],
          },
          sections: [
            {
              properties: {
                page: {
                  margin: {
                    top: docx.convertInchesToTwip(0.8),
                    bottom: docx.convertInchesToTwip(0.5),
                    right: docx.convertInchesToTwip(0.38),
                    left: docx.convertInchesToTwip(0.7),
                  }
                },
              },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Custom Millwork",
                      size: 48,
                      bold: true,
                    }),
                  ], 
                  alignment: 'center'
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "1200 Murphy Drive",
                      size: 32,
                      bold: true,
                    }),
                  ],
                  alignment: 'center'
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Maumelle, AR 72113",
                      size: 32,
                      bold: true,
                    }),
                  ],
                  alignment: 'center'
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Phone: 501.851.4421",
                      size: 32,
                      bold: true,
                    }),
                  ],
                  alignment: 'center'
                }),
                new docx.Paragraph({
                  children: [
                    new docx.ImageRun({
                      type: 'png',
                      data: fs.readFileSync(awiimg),
                      transformation: {
                        width: 100,
                        height: 100,
                      },
                    })
                  ],
                  alignment: "center"
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: details.date,
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `TO:		        ${details.contractor}`, // here
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `PROJECT:  	        ${details.job}`, // here
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `ATTENTION:       Estimating`, // here
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `Please consider the following Price to Furnish & Install the Millwork as shown on Drawings`,
                      size: 24,
                      bold: true
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Pricing Qualifications",
                      size: 28,
                      bold: true,
                      underline: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "AWI Custom Grade",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Drawer Pulls:",
                      size: 22,
                      italics: true,
                    }),
                    new docx.TextRun({
                      text: " 4-inch wire pull",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Door Pulls:",
                      size: 22,
                      italics: true,
                    }),
                    new docx.TextRun({
                      text: " 4-inch wire pull",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Hinges:",
                      size: 22,
                      italics: true,
                    }),
                    new docx.TextRun({
                      text: " Blum concealed self-close",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Drawers:",
                      size: 22,
                      italics: true,
                    }),
                    new docx.TextRun({
                      text: " Blum Metabox system",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Project Specific Exclusions",
                      size: 26,
                      bold: true,
                      underline: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Excludes QCP Certification and Labels",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Excludes Arkansas Wage Rate/LEED/FSC certification labels",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Excludes all Lighting and Electrical",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Excludes all Painting and Finishing",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Excludes All Demolition",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Bid Scope",
                      size: 26,
                      bold: true,
                      underline: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Items/areas not included in the bid scope are EXCLUDED unless otherwise noted",
                      size: 22,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                ...scopelist, // groups from above are added in using the spread operator 
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                // Removed because of the new group pricing stuff
                // but the regex is cool so I won't delete it

                // new docx.Paragraph({
                //   children: [
                //     new docx.TextRun({
                //       text: `Base Price for Furnished and installed......$${details.price.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",")}.00`,
                //       size: 24,
                //       bold: true,
                //     })
                //   ],
                // }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                ...notelist,
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "CONDITIONS:",
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                ...conditionslist,
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "Respectfully submitted:",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "CUSTOM MILLWORK                                          Accepted:                      ", // here
                      size: 24,
                    })
                  ],
                  border: {
                    bottom: {
                      color: 'auto',
                      style: 'single',
                      size: 1,
                      space: 1,
                    }
                  },
                  indent: {
                    right: docx.convertInchesToTwip(1)
                  }
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 18,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "By:                                                                            By:                      ", // here
                      size: 24,
                    })
                  ],
                  border: {
                    bottom: {
                      color: 'auto',
                      style: 'single',
                      size: 1,
                      space: 1,
                    }
                  },
                  indent: {
                    right: docx.convertInchesToTwip(1)
                  }
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `                                              ${settings.info.user ? settings.info.user == 'guest' ? 'cm estimator' : settings.info.user : 'cm estimator'}`, // here
                      size: 18,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 18,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "TITLE:            Project Estimator                              TITLE:", // here
                      size: 24,
                    })
                  ],
                  border: {
                    bottom: {
                      color: 'auto',
                      style: 'single',
                      size: 1,
                      space: 1,
                    }
                  },
                  indent: {
                    right: docx.convertInchesToTwip(1)
                  }
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "UPON ACCEPTANCE OF THIS PROPOSAL, PLEASE SIGN AND RETURN AS SOON AS POSSIBLE.",
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: '(A.) IMPORTANT: The following items are considered "Architectural',
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: 'Woodwork" and do not form a part of the proposal unless noted',
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                ...importantlist1,
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),
                ...importantlist2,
              ]
            },
          ],
        })



        if(generateDoc) { // toggle off the word generation
          docx.Packer.toBuffer(doc).then((buffer) => {
            try {
              fs.writeFileSync(saveLocation, buffer);
            } catch {
              isSuccess = false;
              error('could not save document... make sure the file you are trying to overwrite is closed\n\nplease reselect the csv to try again...', '0001');
              // console.log("could not save document... make sure the file being overwritten is closed - please reselect the csv to try again...");
              // return
            }
          });
        }


        // PDFs
        // uses the 'pdfkit' package
        // can just call methods on the document class to add text
        // don't have to have the whole thing defined in a giant constructor

        // save a second document of the same name + '-totals'
        // show totals for each item in the job
        if(isSuccess && generateTotals) {
          let saveLocation2 = saveLocation.split('.docx')[0] + "-totals.pdf";
          let saveLocation3 = saveLocation.split('.docx')[0] + "-details.pdf";

          const doc2 = new PDFDocument({ size: "LETTER", margins: { top: 25, bottom: 25, left: 50, right: 50 } });
          const doc3 = new PDFDocument({ size: "LETTER", margins: { top: 25, bottom: 25, left: 50, right: 50 } });

          let date = new Date();
          let dateStr = ` ${date.getMonth()+1}/${date.getDate()}/${date.getFullYear()} - ${date.getHours()}:${date.getMinutes()}:${date.getSeconds()}`;
          let groupStr = `number of price groups: [${Object.keys(job.groups).length}]`

          if(!settings.settings.generate_details) {
            doc3.pipe(fs.createWriteStream(saveLocation3));
            doc3.image(iconimg, 480, 20, { width: 60 });
            doc3.fontSize(10);
            doc3.font("Courier-Bold");
            doc3.text(`Job: ${details.job} | Created: ${dateStr}`, { width: 425 });
            doc3.text(`Bid Date: ${details.date}`, { width: 425 });
            doc3.text(groupStr, { width: 425 });
            doc3.text(saveLocation3, { width: 425 });
            doc3.text(" ");
            doc3.text(" ");
          } else {
            doc2.pipe(fs.createWriteStream(saveLocation2));
            doc2.fontSize(10);
            doc2.font("Courier-Bold");
            doc2.text(details.job + dateStr);
            doc2.text(saveLocation2);
          }
          Object.keys(totals.groups).forEach((group) => {
            if(!settings.settings.generate_details) {
              doc3.text(" ");
              doc3.font("Courier-Bold");
              doc3.fontSize(10);
              doc3.text(`${group}`, { underline: true, textIndent: 12 });
              doc3.font("Courier");
            } else {
              doc2.text(" ");
              doc2.font("Courier-Bold");
              doc2.text(`${group}`);
            }
            let data = [];
            Object.keys(totals.groups[group].items).sort().forEach((name, j) => {
              let value = totals.groups[group].items[name];
              let row = [name, value.unit == `ft' in"` ? "LF" : value.unit == `ft` ? "LF" : value.unit, Math.ceil(value.measurement)];
              

              if(!settings.settings.generate_details) {
                let PrimaryRows = [];
                let detailsList = [];
                detailsList.push([
                  { 
                    font: { 
                      src: "./saves/Courier-Bold.afm", 
                      family: "Courier-Bold", 
                      size: 8 
                    }, 
                    backgroundColor: "#000",
                    textColor: "#fff",
                    padding: ['0.5em', '0.25em', '0.25em', '0.25em'],
                    colSpan: 1,
                    text: `${j+1}.`
                  }, 
                  { 
                    font: { 
                      src: "./saves/Courier-Bold.afm", 
                      family: "Courier-Bold", 
                      size: 9 
                    }, 
                    backgroundColor: "#f1f6ff",
                    padding: ['0.5em', '0.25em', '0.25em', '0.25em'],
                    colSpan: 2,
                    text: `${row[0]}`
                  }, 
                  { 
                    font: { 
                      src: "./saves/Courier-Bold.afm", 
                      family: "Courier-Bold", 
                      size: 10 
                    }, 
                    backgroundColor: "#f1f6ff",
                    padding: ['0.5em', '0.25em', '0.25em', '0.25em'],
                    text: row[1]
                  }, 
                  { 
                    font: { 
                      src: "./saves/Courier-Bold.afm", 
                      family: "Courier-Bold", 
                      size: 10 
                    },
                    backgroundColor: "#f1f6ff",
                    padding: ['0.5em', '0.25em', '0.25em', '0.25em'],
                    text: row[2]
                  }]);
                PrimaryRows.push(detailsList.length-1);
                doc3.font("Courier");
                doc3.fontSize(8);
                value.details.forEach((detail, i) => {
                  let dRow;
                  if(detail[2] == 'sqft') {
                    dRow = [
                      {
                        text: ``,
                        backgroundColor: "#000",
                        // textColor: "#fff",
                      },
                      {
                        text: detail[0],
                        backgroundColor: "#dfdfdf",
                        colSpan: 1,
                      },
                      {
                        text: ` ${Math.round((detail[3].measurement + Number.EPSILON) * 100) / 100}' x ${detail[3].depth}"`,
                        backgroundColor: "#dfdfdf",
                        colSpan: 1,
                      },
                      {
                        text: [...new Set(detail[4])].join(" "),
                        backgroundColor: "#dfdfdf",
                      },
                      {
                        text: `${Math.round((detail[1] + Number.EPSILON) * 100) / 100}`,
                        backgroundColor: "#dfdfdf",
                      }
                    ]
                  }
                  else {
                    dRow = [
                      {
                        text: ``,
                        backgroundColor: "#000",
                        // textColor: "#fff",
                      },
                      {
                        text: detail[0],
                        backgroundColor: "#dfdfdf",
                        colSpan: 2,
                      },
                      {
                        text: [...new Set(detail[4])].join(" "),
                        backgroundColor: "#dfdfdf",
                      },
                      {
                        text: `${Math.round((detail[1] + Number.EPSILON) * 100) / 100}`,
                        backgroundColor: "#dfdfdf",
                      }
                    ]
                  }
                  detailsList.push(dRow);
                })
                let nextPrimary = PrimaryRows.shift();
                doc3.table({
                  data: detailsList,
                  columnStyles: (i) => {
                    switch(i) {
                      case 0:
                        return { width: 25 };
                      case 1:
                        return { width: 300 };
                      case 2:
                        return { width: 75 };
                      case 3:
                        return { width: 50 };
                      case 4:
                        return { width: 50 };
                    }
                  },
                  rowStyles: (i) => {
                    if (i == nextPrimary) {
                      nextPrimary = PrimaryRows.shift();
                      return { 
                        height: 16,
                        border: {
                          top: 3
                        }
                      }
                    } else {
                      return { height: 10 } 
                    }
                  }
                });
              } else {
                data.push(row);
              }

            })
            data.sort();
            doc2.font("Courier");
            doc2.table({
              data,
              columnStyles: ["*", 50, 50],
            })
          })
  
          if(!settings.settings.generate_details) {
            doc3.text(" ");
            doc3.text(" ");
            doc3.text(`csv used to generate this data: ${file}`);
            // console.log(stats);
            let createdAt = new Date(stats.birthtime);
            doc3.text(`created: ${createdAt.getDate()}/${createdAt.getMonth()+1}/${createdAt.getFullYear()}-${createdAt.getHours()}:${createdAt.getMinutes()}:${createdAt.getSeconds()}`)
            let modifiedAt = new Date(stats.mtime);
            doc3.text(`modified: ${modifiedAt.getDate()}/${modifiedAt.getMonth()+1}/${modifiedAt.getFullYear()}-${modifiedAt.getHours()}:${modifiedAt.getMinutes()}:${modifiedAt.getSeconds()}`)
            
            doc3.end();
          } else {
            doc2.end();
          }
          if(settings.settings.auto_open_word && !devMode) { // auto opens word after 2 second delay
            setTimeout(() => {
              open.openApp(saveLocation, { app: { name: 'word' } });
            }, 2000)
          }
        }
        callback(isSuccess);
        job = {};
        totals = {};
        rows = [];
        isSuccess = true;
        return
      });
  } else {
    error('invalid file... make sure you selected the correct csv and are saving in a valid location', '0002');
    callback(false);
    console.log("cannot parse csv - no file location");
    return
  }
};

module.exports = { calculateProposal };


// need to find and fix the waiting after selecting a file
// usually the first one (save location) is instant
// but the second (csv) has a pause after selecting that can be 5-6 seconds