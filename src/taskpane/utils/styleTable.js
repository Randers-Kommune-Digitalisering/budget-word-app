import { rydAlt, rydSidehoved, fetchAssets } from "./utils.js";

export async function styleTable(context, table, style) {
    //const table = selection.parentTable;  
    // context.trackedObjects.add(table);

    // Styler tabel som standardtabel
    table.styleBuiltIn = "TableGrid";
    table.headerRowCount = 1;

    const rækker = table.rows;
    rækker.load("items");
    await context.sync();
    console.log(rækker.items);

    for (var i = 0; i < rækker.items.length; i++) {
      var celler = rækker.items[i].cells;
      celler.load("items");
      await context.sync();
      for (var k = 0; k < celler.items.length; k++) {
        celler.items[k].body.styleBuiltIn = "Normal";
      }
    } 

    // Fjerner alle rammer
    var borderLocation = Word.BorderLocation.all;
    var border = table.getBorder(borderLocation);
    await context.sync();
    border.set({ type: "none" });
    await context.sync();

    // Indstiller på hele tabeller
    table.headerRowCount = 1;
    table.font.bold = style.font.bold;
    table.font.size = style.font.size;
    table.font.italics = style.font.italics;

    // Rammer
    for (let b = 0; b < style.border.length; b++) {
      console.log(style.border[b].location);
      var locationKey = style.border[b].location;
      var borderLocation = Word.BorderLocation[locationKey];
      var border = table.getBorder(borderLocation);
      for (const [key, value] of Object.entries(style.border[b].style)) {
        console.log(`${key}: ${value}`);
        var borderProperties = {};  
        borderProperties[key] = value;
        border.set(borderProperties);
      } 
    }
    await context.sync();
    
    // Loop over alle rækker
    // Generel række-styling
    for (var i = 0; i < rækker.items.length; i++) {
      for (var key in style.rows.rowStyle) {
        rækker.items[i][key] = style.rows.rowStyle[key];
        await context.sync();
        console.log(rækker.items[i], style.rows.rowStyle[key]);
      }
    }
    await context.sync();

    /*
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
    */
    await context.sync();
}