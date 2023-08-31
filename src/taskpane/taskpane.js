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
    
    // Indlæser dokumenttype fra UI
    const dokumenttypeUI=document.getElementById("dokumentDropdown").selectedIndex;
    //console.log(dokumenttypeUI);

    // Indlæser dokumenttype parametre fra json
    const response = await fetch('./assets/dokumenttype.json');
    const dokumenttypeJSON = await response.json();
    //console.log(dokumenttypeJSON);

    // Henter kolonneoverskrifter for tabel 1
    const valgtIndex=dokumenttypeUI-1; 
    const dokumenttypeAfgr=dokumenttypeJSON[valgtIndex].tabelindhold;
    //console.log(dokumenttypeAfgr);

    //Udtrækker kolonnenavne for tabel 1
    //for (var key in dokumenttypeAfgr[0].kolonnenavneTabel1) {
    //  context.document.body.insertParagraph(dokumenttypeAfgr[0].kolonnenavneTabel1[key], Word.InsertLocation.end); 
    //}
    const antalKolonner=dokumenttypeAfgr[0].kolonnenavneTabel1.length;
    const kolonneNavne=dokumenttypeAfgr[0].kolonnenavneTabel1;
    //console.log(dokumenttypeAfgr[0].kolonnenavneTabel1.length);

    //Udtrækker delområder
    const udvalgUI=document.getElementById("udvalgDropdown").selectedIndex;
    const bevillingsområdeUI=document.getElementById("bevillingsomrDropdown").selectedIndex;
    
    const responseOrganisation = await fetch('./assets/organisation.json');
    const organisationJSON = await responseOrganisation.json();

    const udvalgIndex=udvalgUI-1;
    const bevillingsområdeIndex=bevillingsområdeUI-1;

    const organisationAfgr= organisationJSON[udvalgIndex].bevillingsomr[bevillingsområdeIndex];
    const delområder=organisationAfgr.delområde;
    const antalRækker=organisationAfgr.delområde.length+1;
    console.log(delområder, antalRækker)

    //const currentYear = new Date(Date.now()).getFullYear();
    //const budgetperiode=[currentYear+1,currentYear+2,currentYear+3,currentYear+4];
    //const overskrift=[""].concat(budgetperiode);

  
    const data = [
      kolonneNavne,
    ];
    const table = context.document.body.insertTable(antalRækker, antalKolonner, "Start", data);
   
    const tabelRækker=table.rows;
    tabelRækker.load('items');

    await context.sync();

    for (var i = 1;i<=tabelRækker.items.length;i++){
      console.log(tabelRækker.items[i].values);
      const test=tabelRækker.items[i].values=[[1,2,3,4,5,6,7,8,9]];
      await context.sync();
    }


    // TODO: Loop gennem rækker og indsæt rækkenavne


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
    header.font.set({
      bold: false,
      italic: false,
      name: "Calibri",
      color: "black",
      size: 18
    });
    
    //header.style.font.size=18;

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