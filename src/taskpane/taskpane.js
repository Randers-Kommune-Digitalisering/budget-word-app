/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("insertTable").onclick = insertTable;
    document.getElementById("addHeader").onclick = addHeader;
  }
});

export async function insertTable() {
  return Word.run(async (context) => {
    // https://www.youtube.com/watch?v=9u6MGqf1J_I
    const currentYear = new Date(Date.now()).getFullYear();
    const budgetperiode=[currentYear+1,currentYear+2,currentYear+3,currentYear+4];
    const overskrift=[""].concat(budgetperiode);
  
    const data = [
      overskrift,
      ["Indtægt", "-3,2", "-", "-", "-"],
      ["Budget", "3,1","0,1","0,1","0,1"],
      ["Nettoresultat", "-0,1","0,1","0,1","0,1"],
    ];
    const table = context.document.body.insertTable(5, 5, "Start", data);
   
    const tablerows=table.rows;

    const row_1=tablerows.getFirst();

    /*
      row_1.horizontalAlignment="Centered";

    row_1.font.bold=true;
*/
    // context.document.body.insertParagraph("test", Word.InsertLocation.end);

    await context.sync();
  });
 
}

export async function addHeader() {
  return Word.run(async (context) => {
    const header1=document.getElementById("udvalgDropdown").value;
    const header2=document.getElementById("bevillingsomrDropdown").value;

    const header=context.document.sections
      .getFirst()
      .getHeader(Word.HeaderFooterType.primary)
      .insertParagraph(header1.concat(" - ", header2), "End");

    header.alignment="Centered";

    await context.sync();
  });
}

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    const response = await fetch('./assets/organisation.json');
    const organisation = await response.json();

    // insert a paragraph at the end of the document.
    for (var key in organisation) {
      if (organisation.hasOwnProperty(key)) {
        for (var key2 in organisation[key].bevillingsomr) {
          if (organisation[key].bevillingsomr.hasOwnProperty(key2)) {
            const tekst=organisation[key].udvalg + " - " + organisation[key].bevillingsomr[key2]
            context.document.body.insertParagraph(tekst, Word.InsertLocation.end); 
          }
        }
      }
    }
    //await context.sync()
    //context.document.save();
    //const paragraph2 = context.document.body.insertParagraph(organisation[1].udvalg, Word.InsertLocation.end);

    // change the paragraph color to blue.
    // paragraph.font.color = "blue";


    await context.sync();
  });
}
/* 
// Function to format table as specified
export async function formatTable() {
  return Word.run(async (context) => {
    // Load the current selection
    var selection = context.document.getSelection();

    // Load the tables in the selection
    var tables = selection.tables;
    context.load(tables);

    // Execute the queued commands
    return context.sync()
      .then(function () {
        // Loop through each table
        for (var i = 0; i < tables.items.length; i++) {
          var table = tables.items[i];
          
          // Set table properties
          table.style.borders.load("items");
          for (var j = 0; j < table.style.borders.items.length; j++) {
            table.style.borders.items[j].color = "#000000"; // Black color
            if (j === 0) {
              // First border is the outer border, set thickness to 2 points
              table.style.borders.items[j].weight = "2pt";
            } else {
              // Inner borders (between cells), set thickness to 0 points to remove them
              table.style.borders.items[j].weight = "0pt";
            }
          }
          
          // Set table header properties
          var tableRows = table.rows;
          context.load(tableRows);
          tableRows.load("items");

          tableRows.items[0].font.bold = true; // Set header rows to bold
          tableRows.items[0].font.color = "#0000FF"; // Blue color for header text
          
          // Set table body properties
          for (var k = 1; k < tableRows.items.length; k++) {
            tableRows.items[k].font.color = "#FFFFFF"; // White color for body text
          }
        }
        
        // Execute the queued commands to update the table formatting
        context.sync();
      });
  });
} */