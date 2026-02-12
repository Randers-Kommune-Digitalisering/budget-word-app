import { rydAlt, rydSidehoved, fetchAssets } from "./utils.js";

function styleCells(cellStyle, cells){
  for (var i = 0; i < cells.items.length; i++) {
    if (i == 0) { 
      for (const [key, value] of Object.entries(cellStyle.left.padding)) {
        cells.items[i].setCellPadding(key, value);
      }
      cells.items[i].horizontalAlignment = cellStyle.left.alignment;
    } else {
      for (const [key, value] of Object.entries(cellStyle.rest.padding)) {
        cells.items[i].setCellPadding(key, value);
      }
      cells.items[i].horizontalAlignment = cellStyle.rest.alignment;
    }
  }
} 

async function styleRow (row, style) {
  for (var key in style) {
    if (key === "font") {
      for (var fontKey in style.font) {
        row.font[fontKey] = style.font[fontKey];
      }
    } 
    else if (key === "border") {
      for (let b = 0; b < style.border.length; b++) {
        var locationKey = style.border[b].location;
        var borderLocation = Word.BorderLocation[locationKey];
        var border = row.getBorder(borderLocation);
        for (const [key, value] of Object.entries(style.border[b].style)) {
          var borderProperties = {};
          borderProperties[key] = value;              
          border.set(borderProperties);
        }
      }
    }
    else if (key === "cells") {
      var cells = row.cells;
      cells.load("items");
      await row.context.sync();

      styleCells(style.cells, cells);
    } 
    else {
      row[key] = style[key];
    }
  }
}

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

    // Tabelbredde
    const sidebredde = style.sidebredde;
    const tabelbredde = sidebredde;
    table.width = sidebredde;
    await context.sync();

    // Rammer
    for (let b = 0; b < style.border.length; b++) {
      var locationKey = style.border[b].location;
      var borderLocation = Word.BorderLocation[locationKey];
      var border = table.getBorder(borderLocation);
      for (const [key, value] of Object.entries(style.border[b].style)) {
        var borderProperties = {};  
        borderProperties[key] = value;
        border.set(borderProperties);
      } 
    }
    await context.sync();

    // Loop over alle rækker
    // Generel række-styling
    for (var i = 0; i < rækker.items.length; i++) {
      // Alle rækker
      for (var key in style.rows.rowStyle) {
        rækker.items[i][key] = style.rows.rowStyle[key];
        await context.sync();
      }

      // Indlæs celler
      var celler = rækker.items[i].cells;
      celler.load("items");
      await context.sync();

      // Indstiller kolonnebredde af første kolonne
      var venstreKolonne = style.breddeVenstreKolonne1;
      if (celler.items.length > style.breddejusteringTreshold) {
        venstreKolonne = style.breddeVenstreKolonne2;
      }
      
      // Indstiller padding
      styleCells(style.cells, celler); 

      for (var k = 0; k < celler.items.length; k++) {     
        // Instiller kolonnebredde
        if (k == 0) {
          celler.items[k].columnWidth = tabelbredde * venstreKolonne;
        }
        if (k >= 1) {
          celler.items[k].columnWidth = tabelbredde * ((1-venstreKolonne)/(celler.items.length-1));
        }
      }

      // Headerrækker
      if (i < style.rows.headerRows) {
        styleRow(rækker.items[i], style.rows.headerRowStyle);
      }
      
      // Foderrækker
      if (i >= rækker.items.length - style.rows.footerRows) {
        styleRow(rækker.items[i], style.rows.footerRowStyle);
      }

      // Custom rækker
      if (style.rows.hasOwnProperty("customRows")) {
        for (let c = 0; c < style.rows.customRows.length; c++) {
          var identifier = style.rows.customRows[c].textIdentifier;
          if (celler.items[0].value.includes(identifier)) {
            styleRow(rækker.items[i], style.rows.customRows[c].customRowStyle);
          }
        }
      }

    }
    await context.sync();
}