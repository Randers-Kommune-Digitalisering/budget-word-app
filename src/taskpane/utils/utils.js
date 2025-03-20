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


// Formater tabeller til budgetbemærninger
export async function formaterTabellerBBSelected() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("parentTable");
    await context.sync();

    const table = selection.parentTable;  

    table.headerRowCount = 1;
    table.font.bold = false;
    table.font.size = 9;

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
        // kolonnebredde
        if (k == 0) {
          celler.items[k].columnWidth = 248;
        }
        if (k >= 1) {
          celler.items[k].columnWidth = 61;
        }

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
        if ((k == 0) & (i > 0) & (i < rækker.items.length - 1)) {
          celler.items[k].setCellPadding("Left", 10);
        }
        // Styler i alt-rækker
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
        // Styler "servicerammen"
        if (
          (celler.items[k].value.slice(0, 13) == "Servicerammen") |
          (celler.items[k].value.slice(0, 22) == "Uden for servicerammen") |
          (celler.items[k].value.slice(0, 12) == "Udgangspunkt")
        ) {
          // console.log(celler.items[k].value.slice(0,5))
          rækker.items[i].shadingColor = "#BDD7EE";
          rækker.items[i].font.bold = true;
          celler.items[k].setCellPadding("Top", 3);
          celler.items[k].setCellPadding("Bottom", 3);
          celler.items[k].setCellPadding("Left", 5);
          var borderLocation = Word.BorderLocation.top;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
          var borderLocation = Word.BorderLocation.bottom;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
        }
        // Styler afsnit i tabellen
        if (
          celler.items[k].value.slice(0, 13) == "Budgetaftale " ||
          celler.items[k].value.slice(0, 9) == "Demografi" ||
          celler.items[k].value.slice(0, 26) == "Budgetaftaler tidligere år" ||
          celler.items[k].value.slice(0, 35) == "Tillægsbevillinger og omplaceringer" ||
          celler.items[k].value.slice(0, 23) == "PL og øvrige ændringer" ||
          celler.items[k].value.slice(0, 27) == "Ændringer fra tidligere år"
        ) {
          rækker.items[i].font.bold = true;
        }
        if (celler.items[k].value.slice(0, 21) == "Ændringer fra budget ") {
          rækker.items[i].font.italic = true;
        }
        await context.sync();
      }
    }
  });
}



// Formater tabeller til budgetbemærninger
export async function formaterTabellerBB(tabel) {
  return Word.run(async (context) => {
    // console.log(tabel);

    const tables = context.document.body.tables;
    tables.load("items");

    await context.sync();

    // tilføjer custom tag til hver tabel
    for (let i = 0; i < tables.items.length; i++) {
      tables.items[i].tag = `tabel-${i}`;
    }

    // Afgrænser til relevant tabel
    const table = tables.items.find((table) => table.tag === tabel);
    await context.sync();

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
        // kolonnebredde
        if (k == 0) {
          celler.items[k].columnWidth = 248;
        }
        if (k >= 1) {
          celler.items[k].columnWidth = 61;
        }

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
        if ((k == 0) & (i > 0) & (i < rækker.items.length - 1)) {
          celler.items[k].setCellPadding("Left", 10);
        }
        // Styler i alt-rækker
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
        // Styler "servicerammen"
        if (
          (celler.items[k].value.slice(0, 13) == "Servicerammen") |
          (celler.items[k].value.slice(0, 22) == "Uden for servicerammen") |
          (celler.items[k].value.slice(0, 12) == "Udgangspunkt")
        ) {
          // console.log(celler.items[k].value.slice(0,5))
          rækker.items[i].shadingColor = "#BDD7EE";
          rækker.items[i].font.bold = true;
          celler.items[k].setCellPadding("Top", 3);
          celler.items[k].setCellPadding("Bottom", 3);
          celler.items[k].setCellPadding("Left", 5);
          var borderLocation = Word.BorderLocation.top;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
          var borderLocation = Word.BorderLocation.bottom;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({ color: "#808080", width: 1, type: "Single" });
        }
        // Styler afsnit i tabellen
        if (
          celler.items[k].value.slice(0, 13) == "Budgetaftale " ||
          celler.items[k].value.slice(0, 9) == "Demografi" ||
          celler.items[k].value.slice(0, 26) == "Budgetaftaler tidligere år" ||
          celler.items[k].value.slice(0, 35) == "Tillægsbevillinger og omplaceringer" ||
          celler.items[k].value.slice(0, 23) == "PL og øvrige ændringer" ||
          celler.items[k].value.slice(0, 27) == "Ændringer fra tidligere år"
        ) {
          rækker.items[i].font.bold = true;
        }
        if (celler.items[k].value.slice(0, 21) == "Ændringer fra budget ") {
          rækker.items[i].font.italic = true;
        }
        await context.sync();
      }
    }
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
