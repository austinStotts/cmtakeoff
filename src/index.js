const { app, BrowserWindow, ipcMain, dialog } = require("electron/main");
let fs = require("fs");
const PDFDocument = require("pdfkit");
const { parse } = require("csv-parse");
const open = require('open');
const docx = require("docx");
let version = '1.0.9';
let generateDoc = true; // true
let generateTotals = true; // true
let isSuccess = true;

let devMode = false;
if(process.env.TERM_PROGRAM == 'vscode') {
  devMode = true;
}

// calculates the unique page labels for all children of an Item
// could maybe be a class method
let calculateRooms = (list) => {
  let pages = [];
  let keys = Object.keys(list);
  keys.forEach((item, i) => {
    list[item].children.forEach((child, j) => { pages.push(child.page) })
  })
  return [...new Set(pages)];
}




// 3/10/26
// these is a bug where the incorrect value is shown on the proposal for topshop/sqft items
// it is showing the sqft value but calling it LF
// need to properly convert back to a linear foot on the proposal




// 10-28-25
// finished adding the logic for the add_page_labels toggle
// still need to actually make the ui toggle for it but it works
// need to add the subject=exclusion stuff
// if an excllusion is found in the list of markups we will add it to the proposal exclusion list
// but also maybe show the exclusions on the details sheet
// remove them from the job object though 

// the comment logic is still being worked on
// i want comments to be a per-item thing but show up on a per room basis

// 




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
// the main exported function
// this is what gets invoked from main when the user clicks 'start'
// all the info the user provided is passed in as well as settings and 2 callback functions - error and callback
// the 'callback' callback is to let the user know the job is done
// the error callback is to let the user know something went wrong - it also accepts an error code but so far these dont mean much
let calculateProposal = (file, details, saveLocation, settings, error, callback) => {
  class Item {
    constructor (name, search, material, unit, groupName, spaceName) {
      this.name = name;
      this.search = search;
      this.material = material;
      this.unit = unit;
      this.groupName = groupName;
      this.spaceName = spaceName;
      this.children = [];
    }

    clone () {
      let newItem = new Item(this.name, this.search, this.material, this.unit, this.groupName, this.spaceName);
      this.children.forEach(child => {
        newItem.addChild(child.clone())
      })
      return newItem;
    }

    addChild(child) {
      if(Array.isArray(child)) {
        this.children.concat(child);
      } else {
        this.children.push(child);
      }
    }

    isTops () {
      let isTop = false;
      ['tops', 'countertops'].concat(settings.settings.top_shop_tool_subject.split(',')).forEach(v => {
        if(this.search == v.trim()) {
          isTop = true;
          this.unit = 'sqft';
        }
      })
      return isTop;
    }

    getUnit (bypass=false) {
      if(bypass) return 'LF';
      if(this.children.length > 0) {
        let firstChild = this.children[0];
        if(firstChild.unit == 'sqft' || firstChild.unit == 'sf') {
          return 'sqft';
        } else if(firstChild.unit == 'Count') {
          return 'X';
        } else {
          return 'LF';
        }
        
      }
    }

    writeLine () { // returns a string to be used in the proposal
      let inLines = this.flattenChildren();
      let outLines = [];
      inLines.forEach(line => {
        let total;
        let unit;
        if(this.isTops()) {
          total = line.measurement;
          unit = this.getUnit(this.isTops());
        } else {
          total = this.calculateTotal();
          unit = this.getUnit(this.isTops());
        }
        if(unit == 'X') {
          outLines.push(`(${Math.ceil(total)}${unit}) ${this.name}`);
        } else {
          outLines.push(`${Math.ceil(total)} ${unit} of ${this.name}`);
        }
      })
      return outLines;
    }

    
    calculateTotal(bypass=false) {
      // add up all children and return the total
      return this.children.reduce((total, current) => {
        return total + current.calculateMeasurement(bypass)
      }, 0);
    }

    flattenChildren () {
      // join same room / same depth / children
      // this is certainly not efficient...
      // I am comparing each child to each other child
      // shouldnt be an issue for this scale of items but worth knowing
      let shortList = Object.values(this.children.reduce((accumulator, current) => {
        current.calculateMeasurement(this.isTops())
        let roomName = `${current.spaceName}-${current.depth}-${settings.settings.show_notes ? current.note : ''}`;
        if(!accumulator[roomName]) {
          accumulator[roomName] = current.clone();
        } else {
          accumulator[roomName].measurement += current.measurement;
        }
        return accumulator;
      }, {}))

      // sort the children so that children with comments get grouped
      // warning does not work if the sorting direction is reversed
      return shortList.sort((a, b) => {
        if(a.note < b.note) return 1; // 1 !CANNOT BE -1!
        if(a.note > b.note) return -1; // -1 !CANNOT BE 1!
        return 0;
      });
    }
  }

  class Child {
    constructor (name, search, measurement, depth, unit, groupName, spaceName, page, note, excluded) {
      this.name = name;
      this.search = search;
      this.measurement = measurement;
      this.depth = depth;
      this.unit = unit;
      this.groupName = groupName;
      this.spaceName = spaceName;
      this.page = page;
      this.note = note;
      this.excluded = excluded;
    }

    clone () {
      return new Child(this.name, this.search, this.measurement, this.depth, this.unit, this.groupName, this.spaceName, this.page, this.note, this.excluded);
    }

    calculateMultiplier () {
      let n = 1;
      let name = this.spaceName.trim();
      let m = name.match(/\((?:\d+X|X\d+)\)/i);
      if(m) {
        n = parseInt(m[0].replace(/\D/g, ''));
      }
      // let roomarray = this.spaceName.trim().split(" ");
      // let n = 1;
      // if (roomarray[0].toLowerCase().includes("x)")) {
      //   n = roomarray[0].slice(1, roomarray[0].length - 2);
      // } else if (roomarray[0].toLowerCase().includes("(x")) {
      //   n = roomarray[0].slice(2, roomarray[0].length - 1);
      // } else if (roomarray[roomarray.length - 1].toLowerCase().includes("x)")) {
      //   n = roomarray[roomarray.length - 1].slice(1, roomarray[roomarray.length - 1].length - 2);
      // } else if (roomarray[roomarray.length - 1].toLowerCase().includes("(x")) {
      //   n = roomarray[roomarray.length - 1].slice(2, roomarray[roomarray.length - 1].length - 1);
      // } else {
      //   // none
      // }
      return n;
    }

    calculateMeasurement (bypass=false) {
      if(this.depth > 0 && bypass) {
        this.unit = 'sqft';
        return (this.measurement * (this.depth / 12)) * this.calculateMultiplier();
      } else {
        return this.measurement * this.calculateMultiplier();
      }
    }
  }

  // log some usefull data
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
    // reads the .csv
    // first line is the column headers 
    fs.createReadStream(file)
      .pipe(parse({ delimiter: ",", from_line: 1 }))
      .on("data", function (row) {
        rows.push(row);
      })
      .on("end", () => {
        isSuccess = true;
        console.log("input successful");
        console.log(settings);
        let columns = rows.shift(); // remove the first index because these are column names and not actual data
        columns[0] = columns[0].slice(1);
        console.log(columns);
        let toolKey = settings.settings.primary_column || "Subject";
        let searchKey = toolKey == "Subject" ? "Label" : "Subject";
        let totals = { groups: {} };
        let job = { groups: {} };
        let exclusions = [];
        for (let i = 0; i < rows.length; i++) {
          let row = {};
          for(let j = 0; j < columns.length; j++) {
            row[columns[j]] = rows[i][j];
          }

          if(row[searchKey] == 'exclusion') {
            exclusions.push(row);
            continue;
          }

          if(!row['Group']) { row['Group'] = 'Base Bid' } // a blank group will be changed to 'Base Bid'
          
          // JOB
          // the following nested if blocks take each row of data from the csv and organize it into group -> space -> item
          // each row will either be a new Item or be added as a child of an existing Item
          // Items can have children so that similar things can be grouped without erasing important unique info
          // for example there may be many 'solid surface tops' items but each might have a different depth
          // we still want to group all the solid surface top but want to be able to calculate sqft on a measurement by measurement basis
          if(job.groups[row['Group']]) {
            if(job.groups[row['Group']].rooms[row['Space']]) {
              if(job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`]) { // old item
                job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`].addChild(new Child(
                  `${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`,
                  row[searchKey],
                  Number(row['Measurement']),
                  row['Depth'] ? row['Depth'] : "",
                  row['Measurement Unit'],
                  row['Group'],
                  row['Space'],
                  row['Page Label'],
                  row['Note'] || '',
                  false,
                ))
              } else { // new item
                job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`] = new Item(
                  `${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`, 
                  row[searchKey], 
                  row['Material'] || '*',
                  row['Measurement Unit'],
                  row['Group'],
                  row['Space'],
                );
                job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`].addChild(new Child(
                  `${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`,
                  row[searchKey],
                  Number(row['Measurement']),
                  row['Depth'] ? row['Depth'] : "",
                  row['Measurement Unit'],
                  row['Group'],
                  row['Space'],
                  row['Page Label'],
                  row['Note'] || '',
                  false,
                ))
              }
            } else { // new space
              job.groups[row['Group']].rooms[row['Space']] = { items: {} };
              job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`] = new Item(
                `${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`, 
                row[searchKey], 
                row['Material'] || '*',
                row['Measurement Unit'],
                row['Group'],
                row['Space'],
              );
              job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`].addChild(new Child(
                `${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`,
                row[searchKey],
                Number(row['Measurement']),
                row['Depth'] ? row['Depth'] : "",
                row['Measurement Unit'],
                row['Group'],
                row['Space'],
                row['Page Label'],
                row['Note'] || '',
                false,
              ))
            }
          } else { // new group
            job.groups[row['Group']] = { rooms: {} };
            job.groups[row['Group']].rooms[row['Space']] = { items: {} };
            job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`] = new Item(
              `${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`, 
              row[searchKey], 
              row['Material'] || '*',
              row['Measurement Unit'],
              row['Group'],
              row['Space'],
            );
            job.groups[row['Group']].rooms[row['Space']].items[`${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`].addChild(new Child(
              `${row[toolKey]}${row['Material']?` (${row['Material']})`:""}`,
              row[searchKey],
              Number(row['Measurement']),
              row['Depth'] ? row['Depth'] : "",
              row['Measurement Unit'],
              row['Group'],
              row['Space'],
              row['Page Label'],
              row['Note'] || '',
              false,
            ))
          }
        }


        // TOTALS
        
        Object.keys(job.groups).forEach(group => {
          if(!totals.groups[group]) {
            totals.groups[group] = { items: {} }
          }
          Object.keys(job.groups[group].rooms).forEach(room => {
            Object.keys(job.groups[group].rooms[room].items).forEach(item => {
              let jobItem = job.groups[group].rooms[room].items[item];
              if(totals.groups[group].items[item]) { // old item
                jobItem.children.forEach(child => {
                  totals.groups[group].items[item].addChild(child.clone())
                })
              } else { // new item
                totals.groups[group].items[item] = jobItem.clone();
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
                text: `${group}`,
                bold: true,
                size: 24,
                underline: true,
              })
            ]
          }))
          Object.keys(job.groups[group].rooms).forEach((room, i) => {
            let items = job.groups[group].rooms[room].items;
            scopelist.push(new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: settings.settings.add_page_labels ? `${room} as seen on [${calculateRooms(items).join(', ')}]` : `${room}`,
                  bold: true,
                  size: 24,
                })
              ],
              numbering: {
                reference: "my-numbering",
                level: 0,
              }
            }));

            Object.keys(job.groups[group].rooms[room].items).forEach(item => {
              let items = job.groups[group].rooms[room].items;
              let lines = items[item].writeLine();
              lines.forEach(line => {
                scopelist.push(new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: line,
                      size: 24,
                    })
                  ],
                  numbering: {
                    reference: "my-bullets",
                    level: 0,
                  }
                }));
              })
            })
          })

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
        

        exclusionsList = [];
        exclusionsText = [
          "Excludes QCP Certification and Labels",
          "Excludes Arkansas Wage Rate/LEED/FSC certification labels",
          "Excludes all Lighting and Electrical",
          "Excludes all Painting and Finishing",
          "Excludes All Demolition",
        ];

        if(settings.settings.include_default_exclusions) {
          exclusionsText.forEach(text => {
            exclusionsList.push(new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: text,
                  size: 22,
                })
              ], 
            }))
          })
        }

      exclusions.forEach(text => {
          exclusionsList.push(new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: text['Comments'],
                size: 22,
              })
            ], 
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
                ...exclusionsList,
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
            }
          });
        }


        // PDFs
        // uses the 'pdfkit' package
        // can just call methods on the document class to add text
        // don't have to have the whole thing defined in a giant constructor

        // save a second document of the same name + '-totals' <- old
        // show totals for each item in the job

        // totals are not being generated but doc2 and doc3 are sort of woven together
        // will remove the doc2 stuff at some point
        // currently all doc2 code is not needed
        if(isSuccess && generateTotals) {
          let saveLocation2 = saveLocation.split('.docx')[0] + "-totals.pdf";
          let saveLocation3 = saveLocation.split('.docx')[0] + "-details.pdf";

          const doc2 = new PDFDocument({ size: "LETTER", margins: { top: 25, bottom: 25, left: 50, right: 50 } });
          const doc3 = new PDFDocument({ size: "LETTER", margins: { top: 25, bottom: 25, left: 50, right: 50 } });

          let date = new Date();
          let dateStr = ` ${date.getMonth()+1}/${date.getDate()}/${date.getFullYear()} - ${date.getHours() > 12 ? date.getHours()-12:date.getHours()}:${date.getMinutes()}:${date.getSeconds()}`;
          let groupStr = `number of price groups: [${Object.keys(job.groups).length}]`

          if(!settings.settings.generate_details) {
            doc3.pipe(fs.createWriteStream(saveLocation3));
            doc3.image(iconimg, 480, 20, { width: 60 });
            doc3.fontSize(10);
            doc3.font("Courier-Bold");
            doc3.text(`Job: ${details.job} | Created: ${dateStr}`, { width: 425 });
            doc3.text(`Bid Date: ${details.date} | To: ${details.contractor}`, { width: 425 });
            doc3.text(groupStr, { width: 425 });
            doc3.fontSize(8);
            doc3.text(saveLocation3, { width: 425 });
            doc3.fontSize(10);
            doc3.text(" ");
            doc3.text(" ");
            doc3.text("Exclusions", { underline: true });
            exclusions.forEach(exclusion => {
              doc3.text(`${exclusion['Comments']}`);
            })
          } else {
            doc2.pipe(fs.createWriteStream(saveLocation2));
            doc2.fontSize(10);
            doc2.font("Courier-Bold");
            doc2.text(details.job + dateStr);
            doc2.text(saveLocation2);
          }
          Object.keys(totals.groups).forEach((group) => { // loop through each group
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

            let currentSubject = "";
            let subjectNumber = 0;

            let subjectColors = ['#adffaa', '#a6eaff', '#ffb4d7', '#ffc8a3', '#d1b4ff', '#faffb7'];

            // the order of items in the details sheet is pretty all over the place
            // I want to sort them by subject first and they by label second

            // this whole thing is such a mess now

            // START
            // loop through each item in the totals object
            Object.keys(totals.groups[group].items)
            .sort((a, b) => { // sort the array
              let aS = totals.groups[group].items[a].search;
              let bS = totals.groups[group].items[b].search;

              let aL = totals.groups[group].items[a].name;
              let bL = totals.groups[group].items[b].name;

              // compare the subjects first and if equal, compare the labels
              return aS.localeCompare(bS) || aL.localeCompare(bL)
            })
            .forEach((name, j) => { // the actual loop starts here - still have to reference every key to the totals object
              let value = totals.groups[group].items[name]; // value = item

              //           0                                                      1                                                                     2                                           
              let row = [name, value.isTops() ? 'sqft' : value.unit == `ft' in"` ? "LF" : value.unit == `ft` ? "LF" : value.unit, Math.ceil(value.calculateTotal(value.isTops()))];
              if(!settings.settings.generate_details) { // primary item row - numbered with larger font
                let PrimaryRows = [];
                let detailsList = [];
                let noteList = [];

                if(value.search != currentSubject) { // break up each subject with a colored line
                  detailsList.push([
                    {
                      border: { right: 0 },
                      font: { 
                        src: "./saves/Courier-Bold.afm", 
                        family: "Courier-Bold", 
                        size: 8 
                      }, 
                      backgroundColor: subjectColors[subjectNumber],
                      text: `${value.search.toUpperCase()}`,
                      colSpan: 3,
                      padding: { left: '0.75em' },
                    },
                    {
                      border: { left: 0 },
                      font: { 
                        src: "./saves/Courier-Bold.afm", 
                        family: "Courier-Bold", 
                        size: 8 
                      }, 
                      backgroundColor: subjectColors[subjectNumber],
                      text: `${value.search.toUpperCase()}`,
                      padding: { right: '0.75em' },
                      colSpan: 3,
                      align: { x: 'right' }
                    },
                  ]);
                  currentSubject = value.search;
                  subjectNumber += 1;
                }

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
                    text: `${j+1}.`,
                    align: { x: 'left', y: 'center' },
                    borderColor: { bottom: '#000', top: '#000' },
                  }, 
                  { 
                    font: { 
                      src: "./saves/Courier-Bold.afm", 
                      family: "Courier-Bold", 
                      size: 9 
                    }, 
                    backgroundColor: "#f1f6ff",
                    padding: ['0.5em', '0.25em', '0.25em', '0.25em'],
                    colSpan: 3,
                    text: `${row[0]}`,
                    align: { x: 'left', y: 'center' },
                    borderColor: { bottom: '#000', top: '#000' },
                  }, 
                  { 
                    font: { 
                      src: "./saves/Courier-Bold.afm", 
                      family: "Courier-Bold", 
                      size: 10 
                    }, 
                    backgroundColor: "#f1f6ff",
                    padding: ['0.5em', '0.25em', '0.25em', '0.25em'],
                    text: row[1],
                    align: { x: 'left', y: 'center' },
                    borderColor: { bottom: '#000', top: '#000' },
                  }, 
                  { 
                    font: { 
                      src: "./saves/Courier-Bold.afm", 
                      family: "Courier-Bold", 
                      size: 10,
                      align: { x: 'left', y: 'center' }, 
                    },
                    backgroundColor: "#f1f6ff",
                    padding: ['0.5em', '0.25em', '0.25em', '0.25em'],
                    text: row[2],
                    align: { x: 'left', y: 'center' },
                    borderColor: { bottom: '#000', top: '#000' },
                  }]);
                PrimaryRows.push(detailsList.length-1);
                doc3.font("Courier");
                doc3.fontSize(8);
                currentNote = null;
                let childList = value.flattenChildren();
                childList.forEach((child, i) => {
                  if(currentNote == null) { currentNote = child.note } 
                  if(child.note != currentNote) {
                    if(currentNote) {
                      let cRow = [
                        {
                          font: { 
                            src: "./saves/Courier-Bold.afm", 
                            family: "Courier-Bold", 
                            size: 8 
                          }, 
                          text: ``,
                          backgroundColor: "#000",
                          textColor: "#fff",
                          colSpan: 1,
                        },
                        {
                          font: { 
                            src: "./saves/Courier-Bold.afm", 
                            family: "Courier-Bold", 
                            size: 8 
                          },
                          border: { top: 0, right: 0, bottom: 1 },
                          text: '',
                          backgroundColor: '#cfcfcf',
                          textColor: '#000',
                          align: { x: 'center', y: 'center' }
                        },
                        { 
                          font: {
                            src: "./saves/Courier-Bold.afm", 
                            family: "Courier-Bold", 
                            size: 7,
                          },
                          text: `${currentNote}`,
                          border: { left: 0, bottom: 1 },
                          backgroundColor: "#cfcfcf",
                          colSpan: 4,
                          align: { x: 'left', y: 'center' },
                        },
                      ];
                      if(settings.settings.show_notes) {
                        detailsList.push(cRow);
                        if(currentNote.length > 75) { noteList.push(detailsList.length-1); }
                        currentNote = child.note;
                      }
                    }
                  }
                  let isNote = child.note ? true : false
                  let dRow;
                  if(!isNote || !settings.settings.show_notes) { // no note
                    if(child.unit == 'sqft') { // 5 col SQFT
                      dRow = [
                        {
                          text: ``,
                          backgroundColor: "#000",
                        },
                        {
                          text: child.spaceName,
                          backgroundColor: "#dfdfdf",
                          colSpan: 2,
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: ` ${Math.trunc(child.depth)}" x ${Math.round((child.measurement + Number.EPSILON) * 100) / 100}'`, // round the measurement to the nearest 2 decimals and remove all decimals from the depth (should always ve a whole number in inches)
                          backgroundColor: "#dfdfdf",
                          colSpan: 1,
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: child.page,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: `${Math.round((child.calculateMeasurement(value.isTops()) + Number.EPSILON) * 100) / 100}`,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        }
                      ]
                    }
                    else { // 4 cols everything else
                      dRow = [
                        {
                          text: ``,
                          backgroundColor: "#000",
                        },
                        {
                          text: child.spaceName,
                          backgroundColor: "#dfdfdf",
                          colSpan: 3,
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: child.page,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: `${Math.round((child.calculateMeasurement(value.isTops()) + Number.EPSILON) * 100) / 100}`,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        }
                      ]
                    }
                  } else { // yes note
                    if(child.unit == 'sqft') { // 5 col SQFT
                      dRow = [
                        {
                          text: ``,
                          backgroundColor: "#000",
                        },
                        {
                          text: '',
                          backgroundColor: '#bdbdbd',
                          align: { x: 'center', y: 'bottom' },
                          border: { bottom: 0, right: 0, top: 0 },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: child.spaceName,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: ` ${Math.trunc(child.depth)}" x ${Math.round((child.measurement + Number.EPSILON) * 100) / 100}'`, // round the measurement to the nearest 2 decimals and remove all decimals from the depth (should always ve a whole number in inches)
                          backgroundColor: "#dfdfdf",
                          colSpan: 1,
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: child.page,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: `${Math.round((child.calculateMeasurement(value.isTops()) + Number.EPSILON) * 100) / 100}`,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        }
                      ]
                    }
                    else { // 4 cols everything else
                      dRow = [
                        {
                          text: ``,
                          backgroundColor: "#000",
                        },
                        {
                          text: '',
                          backgroundColor: '#cfcfcf',
                          align: { x: 'center', y: 'bottom' },
                          border: { bottom: 0, right: 0, top: 0 },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: child.spaceName,
                          backgroundColor: "#dfdfdf",
                          colSpan: 2,
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: child.page,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        },
                        {
                          text: `${Math.round((child.calculateMeasurement(value.isTops()) + Number.EPSILON) * 100) / 100}`,
                          backgroundColor: "#dfdfdf",
                          align: { x: 'left', y: 'center' },
                          borderColor: { bottom: '#000', top: '#000' },
                        }
                      ]
                    }
                  }
                  detailsList.push(dRow);
                  if(childList.length == 1 || i+1 == childList.length) {
                    if(currentNote.length > 0) {
                      let cRow = [
                        {
                          font: { 
                            src: "./saves/Courier-Bold.afm", 
                            family: "Courier-Bold", 
                            size: 8 
                          }, 
                          text: ``,
                          backgroundColor: "#000",
                          textColor: "#fff",
                          colSpan: 1,
                        },
                        {
                          font: { 
                            src: "./saves/Courier-Bold.afm", 
                            family: "Courier-Bold", 
                            size: 8 
                          },
                          border: { top: 0, right: 0, bottom: 1 },
                          text: '',
                          backgroundColor: '#cfcfcf',
                          textColor: '#000',
                          align: { x: 'center', y: 'center' }
                        },
                        { 
                          font: {
                            src: "./saves/Courier-Bold.afm", 
                            family: "Courier-Bold", 
                            size: 7,
                          },
                          text: `${currentNote}`,
                          border: { left: 0, bottom: 1 },
                          backgroundColor: "#cfcfcf",
                          colSpan: 4,
                          align: { x: 'left', y: 'center' },
                        },
                      ];
                      if(settings.settings.show_notes) {
                        detailsList.push(cRow);
                        if(currentNote.length > 75) { noteList.push(detailsList.length-1); }
                      }
                    }
                  }
                })
                let nextPrimary = PrimaryRows.shift();
                let nextNote = noteList.shift();
                doc3.table({
                  data: detailsList,
                  columnStyles: (i) => {
                    switch(i) {
                      case 0:
                        return { width: 25 };
                      case 1:
                        return { width: 15 };
                      case 2:
                        return { width: 285 };
                      case 3:
                        return { width: 90 };
                      case 4:
                        return { width: 60 };
                      case 5:
                        return { width: 40 };
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
                    } else if (i == nextNote) {
                      nextNote = noteList.shift();
                      return { border: { bottom: 0.5, top: 0.5 }, }
                    } else {
                      return { 
                        height: 12,
                        border: { bottom: 0.5, top: 0.5 },
                      } 
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
            doc3.fontSize(6);
            doc3.text(`csv used to generate this data: ${file}`);
            let createdAt = new Date(stats.birthtime);
            let modifiedAt = new Date(stats.mtime);
            doc3.text(`created: ${createdAt.getMonth()+1}/${createdAt.getDate()}/${createdAt.getFullYear()}-${createdAt.getHours()}:${createdAt.getMinutes()}:${createdAt.getSeconds()} | modified: ${modifiedAt.getMonth()+1}/${modifiedAt.getDate()}/${modifiedAt.getFullYear()}-${modifiedAt.getHours()}:${modifiedAt.getMinutes()}:${modifiedAt.getSeconds()}`)
            let timeDif = date.getTime() - createdAt.getTime();
            doc3.text(`s:${Math.round((timeDif/1000 + Number.EPSILON) * 100) / 100}m:${Math.round((timeDif/(1000*60) + Number.EPSILON) * 100) / 100}h:${Math.round((timeDif/(1000*60*60) + Number.EPSILON) * 100) / 100}`);
            doc3.text(`version: ${version}`)
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