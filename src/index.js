const { app, BrowserWindow, ipcMain, dialog } = require("electron/main");
let fs = require("fs");
const PDFDocument = require("pdfkit");
const { parse } = require("csv-parse");
const open = require('open');
const docx = require("docx");

let generateDoc = true;
let generateTotals = true;
let isSuccess = true;

let devMode = false;
if(process.env.TERM_PROGRAM == 'vscode') {
  devMode = true;
}


let randomsort = (a, b) => 0.5 - Math.random();

//                   array       string
let checkPages = (currentPages, page) => { // old
  let hasPage = false;
  for(let i = 0; i < currentPages.length; i++) {
    if(currentPages[i] == page) { hasPage = true }
  }
  return hasPage;
}

function printValues(obj) { // old
    for (var key in obj) {
        if (typeof obj[key] === "object") {
            printValues(obj[key]);   
        } else {
            console.log(obj[key]);    
        }
    }
}

let formatMeasurement = (row) => { // old
  if(row['Depth'] > 0.01) {
    return { measurement: Number(row['Measurement']) * (row['Depth Unit'] == 'in' ? Number(row['Depth']) / 12 : Number(row['Depth'])), unit: 'sqft' };
  } else {
    return { measurement: Number(row['Measurement']), unit: row['Measurement Unit'] };
  }
}

// get the proposal generator to use the new job object
// change the front end to have clearer instructions

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
        let toolKey = 'Subject';
        let totals = { groups: {} };
        let job = { groups: {} };
        // let scope = {};
        // let totals = {};
        // let groups = new Map;
        // This loop organizes the raw csv data into rooms and items
        // it is mandatory for the csv columns to be in the correct order
        // scope - measurement - label - subject - special comments
        for (let i = 0; i < rows.length; i++) {
          let row = {};
          for(let j = 0; j < columns.length; j++) {
            row[columns[j]] = rows[i][j];
          }
          // if a material is given add that to the item key
          // scope
          // console.log([Number(row['Measurement']), row['Measurement Unit']]);
          // if(scope[row['Space']]) { // room exists
          //   if(scope[row['Space']]['items'][`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`]) { // room and item exist
          //     // console.log(scope[row['Space']]['items'][`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].measurement);
          //     scope[row['Space']]['items'][`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].measurement = scope[row['Space']]['items'][`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].measurement + Number(row['Measurement']);
          //     if(row['Page Label'] && !checkPages(scope[row['Space']]['pages'], row['Page Label'])) { scope[row['Space']]['pages'].push(row['Page Label']) }
          //   } else { // room exists but item does not exist
          //     scope[row['Space']]['items'][`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`] = { measurement: Number(row['Measurement']), unit: row['Measurement Unit'] };
          //     if(row['Page Label'] && !checkPages(scope[row['Space']]['pages'], row['Page Label'])) { scope[row['Space']]['pages'].push(row['Page Label']) }
          //   }
          // } else { // room does not exist
          //   scope[row['Space']] = {items: {}, pages: []};
          //   scope[row['Space']]['items'][`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`] = { measurement: Number(row['Measurement']), unit: row['Measurement Unit'] };
          //   if(row['Page Label'] && !checkPages(scope[row['Space']]['pages'], row['Page Label'])) { scope[row['Space']]['pages'].push(row['Page Label']) }
          // }

          if(!row['Group']) { row['Group'] = 'Base Bid' }

          // if(job.groups[row['Price Grouping']]) { // does the Price Grouping already exist?
          //   job.groups[row['Price Grouping']].scope.pages.push(row['Page Label']);
          //   job.groups[row['Price Grouping']].scope.items.push({
          //     name: row[toolKey],
          //     measurement: Number(row['Measurement']),
          //     unit: row['Measurement Unit'],
          //     material: row['Material Type'],
          //     depth: row['Depth'] || null,
          //     groupName: row['Price Grouping'],
          //     spaceName: row['Space'],
          //     excluded: false
          //   })
          // } else {
          //   job.groups[row['Price Grouping']] = {
          //     name: row['Price Grouping'], 
          //     scope: {
          //       rooms: {}
          //     }
          //   }
          // }

          if(job.groups[row['Group']]) {
            if(job.groups[row['Group']].rooms[row['Space']]) {
              if(job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`]) { // old item
                job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`].measurement += Number(row['Measurement']);
              } else { // new item
                job.groups[row['Group']].rooms[row['Space']].pages.push(row['Page Label']);
                job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`] = {
                  name: row[toolKey],
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
              measurement: Number(row['Measurement']),
              unit: row['Measurement Unit'],
              material: row['Material'] || '*',
              depth: Number(row['Depth']) || null,
              groupName: row['Group'],
              spaceName: row['Space'],
              excluded: false
            }
          }


          // need to go back and add spaces
          // skipped past them the first time
          // job > groups > spaces > items
          // the totals can just ignore the space


          

          // console.log(totals);
          

          
          // console.log(JSON.stringify(job))

          // printValues(job)
          // totals


          // if(totals[`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`]) {
          //   if(row['Price Grouping']) {
          //     if(row['Price Grouping'] == totals[`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].group) {
          //       totals[`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].measurement = totals[`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].measurement + (formatMeasurement(row).measurement * n);
          //     } else {

          //     }
          //   } else {
          //     totals[`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].measurement = totals[`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`].measurement + (formatMeasurement(row).measurement * n);
          //   }
          // } else {
          //   totals[`${row[toolKey]}${row['Material Type'] ? ` (${row['Material Type']})` : ""}`] = { measurement: formatMeasurement(row).measurement * n, unit: formatMeasurement(row).unit, group: row['Price Grouping'] };
          // }
          // groups.set(row['Price Grouping']);

          // okay so this needs to be reworked again
          // using the key is not enough
          // the group should not change the item but the material should
          // so maybe it doesn't need to be totally reworked
          //
          // 

          // new job object shape
          // check group first to sort
          // give the job object to the totals and use the already group object
          // let job = {
          //   groups: {
          //     "Base Bid": {
          //         rooms: {
          //           "breakroom": {
          //             items: [
          //               {
          //                 name: 'base cabinet',
          //                 measurement: 8,
          //                 unit: 'lf',
          //                 excluded: false,
          //                 groupName: 'Base Bid',
          //                 depth: null,
          //               },
          //               {
          //                 name: 'base cabinet',
          //                 measurement: 8,
          //                 unit: 'lf',
          //                 excluded: false,
          //                 groupName: 'Base Bid',
          //                 depth: null,
          //               },
          //               {
          //                 name: 'base cabinet',
          //                 measurement: 8,
          //                 unit: 'lf',
          //                 excluded: false,
          //                 groupName: 'Base Bid',
          //                 depth: null,
          //               },
          //               {
          //                 name: 'base cabinet',
          //                 measurement: 8,
          //                 unit: 'lf',
          //                 excluded: false,
          //                 groupName: 'Base Bid',
          //                 depth: null,
          //               },
          //             ],
          //             pages: [A101],
          //           },
          //         },
                
          //     }
          //   }
          // }


        }

          Object.keys(job.groups).forEach(group => {
            // each group
            if(!totals.groups[group]) {
              totals.groups[group] = { items: {} }
            }

            Object.keys(job.groups[group].rooms).forEach(room => {
              // each room
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
                // each item
                // console.log(job.groups[group].rooms[room].items[item].depth)
                if(totals.groups[group].items[item]) { // old item
                  // console.log("MATCH!");
                  if(job.groups[group].rooms[room].items[item].depth) {
                    totals.groups[group].items[item].measurement += (job.groups[group].rooms[room].items[item].measurement * (job.groups[group].rooms[room].items[item].depth / 12)) * n;
                  } else {
                    totals.groups[group].items[item].measurement += job.groups[group].rooms[room].items[item].measurement * n;
                  }
                } else { // new item
                  if(job.groups[group].rooms[room].items[item].depth) { // depth should be in inches
                    totals.groups[group].items[item] = { measurement: (job.groups[group].rooms[room].items[item].measurement * (job.groups[group].rooms[room].items[item].depth / 12)) * n, unit: 'sqft' }
                  } else {
                    totals.groups[group].items[item] = { measurement: job.groups[group].rooms[room].items[item].measurement * n, unit: job.groups[group].rooms[room].items[item].unit }
                  }
                }
              })
            }) 

          })

          // console.log(totals);

          // Object.keys(totals.groups).forEach(key => {
          //   console.log('\n*** *** *** *** *** *** *** ***', key);
          //   Object.keys(totals.groups[key].items).forEach(item => {
          //     console.log(item, ':', totals.groups[key].items[item].measurement)
          //   })
          // })

          // Object.keys(job.groups).forEach(key => {
          //   console.log('\n*** *** *** *** *** *** *** ***', key);
          //   Object.keys(job.groups[key]).forEach(item => {
          //     console.log(item, ':', job.groups[key][item])
          //   })
          // })

        // before the scope is used in the document all numbers are rounded up
        // roundall(scope);
        // console.log(scope);
        // let rooms = Object.keys(scope);
        
        // rooms.forEach(room => {
        //   let items = Object.keys(scope[room].items);
        //   items.forEach(item => {
        //     console.log(item, scope[room].items[item]);
        //   })
        // })
        // console.log(scope);
        // console.log(totals);
        // console.log(groups)

        // Object.keys(job.groups).forEach(key => {
        //   console.log('\n\n');
        //   console.log(key);
        //   console.log(job.groups[key].scope.items);
        // })

        let awiimg = '';
        if(devMode) {
          awiimg = "./images/AWICQW.png";
        } else {
          awiimg = process.resourcesPath + '/images/AWICQW.png'
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
          for(let room = 0; room < rooms.length; room++) {
            let items = Object.keys(job.groups[group].rooms[rooms[room]].items);
            // console.log(items)

            scopelist.push(new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: `${rooms[room]} ${job.groups[group].rooms[rooms[room]].pages.length > 0 ? 'as seen on (' + job.groups[group].rooms[rooms[room]].pages.join(", ") + ')' : ''}`,
                  bold: true,
                  size: 24,
                })
              ],
              numbering: {
                reference: "my-numbering",
                level: 0,
              }
            }));

            for(let item = 0; item < items.length; item++) {
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
              } else if (job.groups[group].rooms[rooms[room]].items[items[item]].unit == 'sqft') {
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
                  text: `Price to Furnish and Install ${group} From Above:......$`,
                  bold: true,
                  size: 24,
                })
              ],
            })
          )
        })





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
                      text: `TO:		        ${details.contractor}`,
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `PROJECT:  	        ${details.job}`,
                      size: 24,
                      bold: true,
                    })
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `ATTENTION:       Estimating`,
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
                // HERE
                // MAKE n GROUPS AND LIST EACH WITH ITS OWN PRICE
                ...scopelist,
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "",
                      size: 24,
                    })
                  ],
                }),


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
                      text: "CUSTOM MILLWORK                                          Accepted:                      ",
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
                      text: "By:                                                                            By:                      ",
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
                      text: `                                              ${settings.info.user ? settings.info.user == 'guest' ? 'cm estimator' : settings.info.user : 'cm estimator'}`,
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
                      text: "TITLE:            Project Estimator                              TITLE:",
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

        // doc.sections.push({properties: {}, children: [ new docx.Paragraph({ text: "test" }) ]})


        if(generateDoc) {
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



        // save a second document of the same name + '-totals'
        // show totals for each item in the job
        if(isSuccess && generateTotals) {
          let saveLocation2 = saveLocation.split('.docx')[0] + "-totals.pdf";
          const doc2 = new PDFDocument({ size: "LETTER", margins: { top: 25, bottom: 25, left: 50, right: 50 } });
          doc2.pipe(fs.createWriteStream(saveLocation2));
          doc2.fontSize(10);
          doc2.font("Courier-Bold");
          let date = new Date();
          let dateStr = ` ${date.getMonth()+1} / ${date.getDate()} / ${date.getFullYear()}`;
          doc2.text(details.job + dateStr);
          doc2.text(saveLocation2);
          Object.keys(totals.groups).forEach((group) => {
            doc2.text(" ");
            doc2.font("Courier-Bold");
            doc2.text(`${group}`);
            let data = [];
            Object.keys(totals.groups[group].items).forEach((name) => {
              let value = totals.groups[group].items[name];
              // console.log(value)
              
              let row = [name, value.unit == `ft' in"` ? "LF" : value.unit == `ft` ? "LF" : value.unit, Math.ceil(value.measurement)];
              // console.log(row)
              data.push(row);

            })
            data.sort();
            doc2.font("Courier");
            doc2.table({
              data,
              columnStyles: ["*", 50, 50],
            })
          })
  
          doc2.end();
          if(generateDoc) {
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


// totals seems to be working again.
// fix the proposal side now...

// now remove the price and add a start button
// then clean up all the commmented code and organize everything

// still need to clean up
// start button added but needs to update the user on success / failure
