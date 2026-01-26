/* Ryd formatering */
export async function fetchAssets(adr) {
  return Word.run(async (context) => {
    var response = await fetch(adr, { cache: "reload" });
    return response.json();
  });
}

export async function rydSidehoved() {
  return Word.run(async (context) => {
    // Ryd tekst i header
    var header = context.document.sections.getFirst().getHeader("primary");

    var afsnit = header.paragraphs; 
    context.load(afsnit, "text");
    await context.sync();
    for (var i = 0; i < afsnit.items.length; i++) {
      afsnit.items[i].delete();
    }

    // Ryd tekst i header
    var footer = context.document.sections.getFirst().getFooter("primary");

    var afsnit = footer.paragraphs;
    context.load(afsnit, "text");
    await context.sync();
    for (var i = 0; i < afsnit.items.length; i++) {
      afsnit.items[i].delete();
    } 
  });
}

export async function rydAlt() {
  return Word.run(async (context) => {
    // Ryd alt i body
    context.document.body.clear();
    await context.sync();
    console.log(typeof withData);
    if (typeof withData !== 'undefined') {
      withData = false;
    }
  });
} 





/* eslint-disable office-addins/no-context-sync-in-loop */
export function sumArrays(...arrays) {
  const n = arrays.reduce((max, xs) => Math.max(max, xs.length), 0);
  const result = Array.from({ length: n });
  return result.map((_, i) => arrays.map((xs) => xs[i] || 0).reduce((sum, x) => sum + x, 0));
}

function getSelectedRadioValue(name) {
  const radios = document.getElementsByName(name);
  for (let i = 0; i < radios.length; i++) {
    if (radios[i].checked) {
      return radios[i].value;
    }
  }
  return null; // No radio button is selected
}

export async function formatIntermediateSumRow() {
  return Word.run(async (context) => {
    // Henter valg fra UI
    const rowItalics = (document.getElementById("checkboxRow0").checked ? true : false);
    const rowBold = (document.getElementById("checkboxRow1").checked ? true : false);
    const rowSingleLeftPadding = (document.getElementById("checkboxRow2").checked ? true : false);
    const rowDoubleLeftPadding = (document.getElementById("checkboxRow3").checked ? true : false);
    const rowTopBorder = (document.getElementById("checkboxRow40").checked ? true : false);
    const rowBottomBorder = (document.getElementById("checkboxRow41").checked ? true : false);
    const rowLightShading = (document.getElementById("checkboxRow5").checked ? true : false);
    const rowDarkShading = (document.getElementById("checkboxRow6").checked ? true : false);
    const rowsToEditCount = parseInt(document.getElementById("selectRow0").value);
    const rowTopRow = (document.getElementById("checkboxRow7").checked ? true : false);

    // Loader parentTabel og parentTableCell
    const selection = context.document.getSelection();
    selection.load("parentTable, parentTableCell");
    await context.sync();

    const rowToEdit = selection.parentTableCell._R;
    const table = selection.parentTable;  

    const rowsToEdit = []
    for (let l = rowToEdit; l < rowsToEditCount+rowToEdit; l++) {
      // console.log(l)
      rowsToEdit.push(l)
    }
    
    // Loader alle tabelrækker
    const rows = table.rows;
    rows.load("items");
    await context.sync();

    // Styler valgt række
    for (var i = 0; i < rows.items.length; i++) {
      if (rowsToEdit.includes(i)) {
        // Shading 
        if (rowLightShading) {
          rows.items[i].shadingColor = "#DDEBF7";
        } else if (rowDarkShading) {  
          rows.items[i].shadingColor = "#BDD7EE";
        } else {   
          rows.items[i].shadingColor ="#FFFFFF";
        }  
        // Font italics
        if (rowItalics) {
          rows.items[i].font.italic = true;
        } else {
          rows.items[i].font.italic = false;  
        }
        // Font bold
        if (rowBold) {
          rows.items[i].font.bold = true;
        } else {
          rows.items[i].font.bold = false;
        }
        // Border
        if (rowTopBorder) {
          var borderLocation = Word.BorderLocation.top;
          var border = rows.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
        } else {
          var borderLocation = Word.BorderLocation.top;
          var border = rows.items[i].getBorder(borderLocation);
          border.set({ color: "#D9D9D9", width: 0.5, type: "Single" });
        }
        if (rowBottomBorder) {
          var borderLocation = Word.BorderLocation.bottom;
          var border = rows.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
        } else {
          var borderLocation = Word.BorderLocation.bottom;
          var border = rows.items[i].getBorder(borderLocation);
          border.set({ color: "#D9D9D9", width: 0.5, type: "Single" });
        }
        // Padding
        var celler = rows.items[i].cells;
        celler.load("items");
        await context.sync();
        if (rowSingleLeftPadding | rowDoubleLeftPadding) {
          if (rowSingleLeftPadding) {
              celler.items[0].setCellPadding("Left", 10);
          } else if (rowDoubleLeftPadding) {
              celler.items[0].setCellPadding("Left", 20);
          }
        } else {
          celler.items[0].setCellPadding("Left", 5.4);
        }
        if (rowTopRow) {
          celler.items[0].setCellPadding("Top", 10);
          celler.items[0].setCellPadding("Bottom", 10);
        } else {
          celler.items[0].setCellPadding("Top", 2);
          celler.items[0].setCellPadding("Bottom", 2);
        }
      }
    }
  });
}

export async function formatSelectedTableBuildIn() {
  return Word.run(async (context) => {
    // Loader valgt tabel
    const selection = context.document.getSelection();
    selection.load("parentTable");
    await context.sync();
    const table = selection.parentTable;  

    // Styler tabel som standardtabel
    table.styleBuiltIn = "TableGrid";

    // Styler tekst i tabel
    // Loop over alle rækker
    const rækker = table.rows;
    rækker.load("items");
    await context.sync();

    for (var i = 0; i < rækker.items.length; i++) {
      var celler = rækker.items[i].cells;
      celler.load("items");
      await context.sync();
      for (var k = 0; k < celler.items.length; k++) {
        // console.log(celler.items[k]);
        celler.items[k].body.styleBuiltIn = "Normal";
      }
    } 
  });
}


export async function formatSelectedTable() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("parentTable");
    console.log(selection);
    await context.sync();

    const table = selection.parentTable;  
    // context.trackedObjects.add(table);

    // Styler tabel som standardtabel
    table.styleBuiltIn = "TableGrid";
    table.headerRowCount = 1;

    // Styler tekst i tabel
    // Loop over alle rækker
    const rækker = table.rows;
    rækker.load("items");
    await context.sync();

    for (var i = 0; i < rækker.items.length; i++) {
      var celler = rækker.items[i].cells;
      celler.load("items");
      await context.sync();
      for (var k = 0; k < celler.items.length; k++) {
        celler.items[k].body.styleBuiltIn = "Normal";
      }
    } 

    table.headerRowCount = 1;
    table.font.bold = false;
    table.font.size = 9;

    // Fjerner alle rammer
    var borderLocation = Word.BorderLocation.all;
    var border = table.getBorder(borderLocation);
    await context.sync();
    border.set({ type: "none" });

    // Tilføjer horisontale streger
    var borderLocation = Word.BorderLocation.insideHorizontal;
    var border = table.getBorder(borderLocation);
    await context.sync();
    border.set({ color: "#D9D9D9", width: 0.5, type: "Single" });

    // Loop over alle rækker

    for (var i = 0; i < rækker.items.length; i++) {
      rækker.items[i].verticalAlignment = "Center";
      rækker.items[i].preferredHeight = 0;

      // Styler første og sidste række
      if ((i == 0) | (i == rækker.items.length - 1) & (document.getElementById("checkbox1").checked)) {
        var borderLocation = Word.BorderLocation.bottom;
        var border = rækker.items[i].getBorder(borderLocation);
        border.set({ color: "#808080", width: 1, type: "Single" });
        var borderLocation = Word.BorderLocation.top;
        var border = rækker.items[i].getBorder(borderLocation);
        border.set({ color: "#808080", width: 1, type: "Single" });
        rækker.items[i].shadingColor = "#DDEBF7";
        rækker.items[i].font.bold = true;
        rækker.items[i].font.name = "Calibri";
      }

      if ( (i == rækker.items.length - 1) & (document.getElementById("checkbox1").checked==false)) {
        var borderLocation = Word.BorderLocation.bottom;
        var border = rækker.items[i].getBorder(borderLocation);
        border.set({ color: "#808080", width: 1, type: "Single" });
      }

      // Loop over celler
      var celler = rækker.items[i].cells;
      celler.load("items");
      await context.sync();

      const sidebredde = 468;
      const tabelBredde = sidebredde;
      var venstreKolonne = 0.4 
      if (celler.items.length > 3) {
        venstreKolonne = 0.3;
      }
      
      table.width = sidebredde;

      for (var k = 0; k < celler.items.length; k++) {     
        // Instiller kolonnebredde
        if (k == 0) {
          celler.items[k].columnWidth = tabelBredde * venstreKolonne;
        }
        if (k >= 1) {
          celler.items[k].columnWidth = tabelBredde * ((1-venstreKolonne)/(celler.items.length-1));
        }

        // Sætter padding
        if (i == 0) {
          celler.items[k].setCellPadding("Top", 10);
          celler.items[k].setCellPadding("Bottom", 10);
        } else  {
          celler.items[k].setCellPadding("Top", 2);
          celler.items[k].setCellPadding("Bottom", 2);
        }
         
        // Tekstfarve
        celler.items[k].body.font.set({color: "#000000"});

        // Højrestiller kolonneoverskrifter til højre hvis kun tal -
        // OBS: Redundant - gælder åbenbart ikke ikke BB-tabeller!
        if ((i == 0) & (k > 0)) {
          const årstal = /^\d{4}$/;
          if (årstal.test(celler.items[k].value) == true) {
            celler.items[k].horizontalAlignment = "Centered";
          } else {
            celler.items[k].horizontalAlignment = "Centered";
          }
        }
        // Højrestiller kolonne > 1 og række > 1
        if ((i > 0) & (k > 0)) {
          celler.items[k].horizontalAlignment = "Right";
        }
        if (document.getElementById("checkbox2").checked) {
          if ((k == 0) & (i > 0) & (i < rækker.items.length - 1)) {
            celler.items[k].setCellPadding("Left", 10);
          }
        }
        // Styler i alt-rækker
        if (celler.items[k].value.slice(0, 5) == "I alt") {
          // console.log(celler.items[k].value.slice(0, 5));
          rækker.items[i].shadingColor = "#DDEBF7";
          rækker.items[i].font.bold = true;
          celler.items[k].setCellPadding("Left", 5);
          var borderLocation = Word.BorderLocation.top; 
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
          var borderLocation = Word.BorderLocation.bottom;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
        }
        
      }
      await context.sync();
    }
    await context.sync();
  });
}



// Formatter tabeller
export async function formaterTabeller() {
  return Word.run(async (context) => {
    // Loop over alle tabeller
    var tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    if (tables.items.length > 0) {
      for (var j = 0; j < tables.items.length; j++) {
        var table = tables.items[j];
        table.headerRowCount = 1;
        table.font.bold = false;
        table.font.size = 8;

        // Fjerner alle rammer
        var borderLocation = Word.BorderLocation.all;
        var border = table.getBorder(borderLocation);
        border.set({ type: "none" });

        // Tilføjer horisontale streger
        var borderLocation = Word.BorderLocation.insideHorizontal;
        var border = table.getBorder(borderLocation);
        border.set({ color: "#D9D9D9", width: 1, type: "Single" });

        // Loop over alle rækker
        var rækker = table.rows;
        rækker.load("items");
        await context.sync();
        for (var i = 0; i < rækker.items.length; i++) {
          rækker.items[i].verticalAlignment = "Center";

          // Styler førster og sidste række
          if ((i == 0) | (i == rækker.items.length - 1)) {
            var borderLocation = Word.BorderLocation.top;
            var border = rækker.items[i].getBorder(borderLocation);
            border.set({ color: "#808080", width: 1, type: "Single" });
            var borderLocation = Word.BorderLocation.bottom;
            var border = rækker.items[i].getBorder(borderLocation);
            border.set({ color: "#808080", width: 1, type: "Single" });
            rækker.items[i].shadingColor = "#DDEBF7";
            rækker.items[i].font.bold = true;
            rækker.items[i].font.name = "Calibri";
          }

          // Loop over celler
          var celler = rækker.items[i].cells;
          celler.load("items");
          await context.sync();
          for (var k = 0; k < celler.items.length; k++) {
            // Sætter padding
            celler.items[k].setCellPadding("Top", 3);
            celler.items[k].setCellPadding("Bottom", 3);

            // Højrestiller kolonneoverskrifter til højre hvis kun tal
            if (i == 0) {
              const årstal = /^\d{4}$/;
              if (årstal.test(celler.items[k].value) == true) {
                celler.items[k].horizontalAlignment = "Right";
              }
            }
            // Højrestiller kolonne > 1 og række > 1
            if ((i > 0) & (k > 0)) {
              celler.items[k].horizontalAlignment = "Right";
            }
            if (k == 0) {
              celler.items[k].columnWidth = 240; // Virker vist ikke efter hensigten...
            }
            if (celler.items[k].value.slice(0, 5) == "I alt") {
              // console.log(celler.items[k].value.slice(0, 5));
              rækker.items[i].shadingColor = "#DDEBF7";
              rækker.items[i].font.bold = true;
              celler.items[k].setCellPadding("Top", 3);
              celler.items[k].setCellPadding("Bottom", 3);
              var borderLocation = Word.BorderLocation.top;
              var border = rækker.items[i].getBorder(borderLocation);
              border.set({ color: "#808080", width: 1, type: "Single" });
              var borderLocation = Word.BorderLocation.bottom;
              var border = rækker.items[i].getBorder(borderLocation);
              border.set({ color: "#808080", width: 1, type: "Single" });
            }
          }
        }
      }
      await context.sync();
    }
    await context.sync();
  });
}


export async function replaceWordsWithLinks() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const text = body.text;
    const regex = /(http(s)?:\/\/.)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/gi;
    const matches = text.match(regex);

    if (matches) {
      for (const match of matches) {
        const searchResults = body.search(match, { matchCase: false, matchWholeWord: true });
        context.load(searchResults, "items");
        await context.sync();

        for (const result of searchResults.items) {
          result.insertHyperlink(match, match, "DisplayText");
        }
      }
    }
  }).catch(function (error) {
    console.log(error.message);
  });
}

// New function to format tables - added as to not change the original formaterSelectedTable function
// It has been rewritten - did not really follow the other - tried to make it more readable
export async function styleTable(table) {
  // Settinng some default values for the table - STARTS HERE
  // Style entire table 
  table.styleBuiltIn = "TableGrid";
  table.horizontalAlignment = "Centered";
  table.verticalAlignment = "Center";
  table.headerRowCount = 1;
  table.font.name = "Calibri";
  table.font.color = "#000000";
  table.font.bold = false;
  table.font.size = 9;
  table.width = 468;
  table.cellPadding = 0;
  table.setCellPadding("Top", 2);
  table.setCellPadding("Bottom", 2);
  
  // Remove all borders
  var borderLocation = Word.BorderLocation.all;
  var border = table.getBorder(borderLocation);
  await table.context.sync();
  border.set({ type: "none" });

  // Add horizontal lines
  var borderLocation = Word.BorderLocation.insideHorizontal;
  var border = table.getBorder(borderLocation);
  await table.context.sync();
  border.set({ color: "#D9D9D9", width: 0.5, type: "Single" });
  // Settinng some default values for the table - ENDS HERE

  // get rows
  const rows = table.rows;
  rows.load("items");
  await table.context.sync();

  // Setting styles unique to the first column and getting which rows have special styling - STARTS HERE
  // Find rows with special styling and add unique styling for first column
  var rowsWithSpecialStyling = [];
  var rowIndex = 0;

  for (const row of rows.items) {
    const firstCell = row.cells.getFirst();
    firstCell.load("value, horizontalAlignment, cellPadding");
    await table.context.sync();

    // get rows with special styling
    if (firstCell.value == 'I alt' || rowIndex === 0) {
      firstCell.setCellPadding("Left", 5);  // set padding for first cell
      rowsWithSpecialStyling.push(rowIndex);
    } else {
      firstCell.setCellPadding("Left", 10); // set padding for first cell
    }

    // make first column left aligned
    firstCell.horizontalAlignment = "Left";
    rowIndex++;
  }
  // Setting styles unique to the first column and getting which rows have special styling - ENDS HERE

  // Apply styling to rows with special styling - STARTS HERE
  for (var i = 0; i < rows.items.length; i++) {
    if (rowsWithSpecialStyling.includes(i)) {
      // for both header and I alt rows
      var border = rows.items[i].getBorder(Word.BorderLocation.bottom);
      border.set({ color: "#808080", width: 1, type: "Single" });
      var border = rows.items[i].getBorder(Word.BorderLocation.top);
      border.set({ color: "#808080", width: 1, type: "Single" });
      rows.items[i].shadingColor = "#DDEBF7";
      rows.items[i].font.bold = true;
      rows.items[i].font.name = "Calibri";
      }

      // Unique to header - Set padding and column width for header
      // here is also some style applying to entire table (the column width applies to all rows)
      if (i == 0) {
        const row = rows.items[i];
        row.load("cells/items");
        await rows.context.sync();
        const cells = rows.items[i].cells;
        const leftColumnFactor = cells.items.length > 3 ? 0.3 : 0.4;
        for (var j = 0; j < cells.items.length; j++) {
          cells.items[j].setCellPadding("Top", 10);
          cells.items[j].setCellPadding("Bottom", 10);
          if (j == 0) {
            cells.items[j].columnWidth = table.width * leftColumnFactor;
          } else {
            cells.items[j].columnWidth = table.width  * ((1-leftColumnFactor)/(cells.items.length-1));
          }
        }
        await cells.context.sync();
      }
    }
  // Apply styling to rows with special styling - ENDS HERE

  await rows.context.sync();
}