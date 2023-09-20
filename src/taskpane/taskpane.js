/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { writeFileAsync } from "xlsx";

console.log(Office.context.contentLanguage);

var genContentControls=[]


/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("skabelon").onclick = skabelon;
    document.getElementById("insertTable").onclick = insertTable;
    document.getElementById("addHeader").onclick = addHeader;
    document.getElementById("loadContentControls").onclick = loadContentControls;
    document.getElementById("indsætTest").onclick = indsætTest;
    document.getElementById("rydAlt").onclick = rydAlt;
    
  }
});
  
export async function test() {
  return Word.run(async (context) => {
    var titel=context.document.body.insertParagraph("test", Word.InsertLocation.start)
  });
}

export async function rydAlt() {
  return Word.run(async (context) => {
    context.document.body.clear(); 

    await context.sync();
  });
}

export async function indsætConcentControl(name) {
  return Word.run(async (context) => {
    context.document.body.paragraphs.getLast().select("End")
    var selection = context.document.getSelection()
    const cc=selection.insertContentControl("RichText");
    cc.title=name

    genContentControls.push(name)
    await context.sync();
  })
}

export async function indsætOverskrift(tekst) {
  return Word.run(async (context) => {
    context.document.body.paragraphs.getLast().select("End")
    var selection = context.document.getSelection()

    
    const overskrift=selection.insertParagraph(tekst);
    overskrift.font.bold=true;
    overskrift.font.size=14;

    await context.sync();
    await indsætConcentControl(tekst)
    
    selection.insertParagraph('', "After");   

  })
}

export async function indsætUndersektioner(sektion, undersektioner, bevillingsområde) {
  return Word.run(async (context) => {
    context.document.body.paragraphs.getLast().select("End")
    var selection = context.document.getSelection()
    var lc=sektion.toLowerCase()
    for (var key2 in undersektioner) {
      if (undersektioner.hasOwnProperty(key2)) {
        if (undersektioner[key2].hasOwnProperty(lc)) {
        const undersektion=eval(undersektioner[key2][lc])
          for (var key3 in undersektion) {
            if (undersektion.hasOwnProperty(key3)) {
              const tekstUndersektion=undersektion[key3]
              console.log(tekstUndersektion)
                var underoverskrift=selection.insertParagraph(tekstUndersektion)
                underoverskrift.font.italic=true
                context.document.body.paragraphs.getLast().select("End")
                await context.sync();
                await indsætConcentControl(sektion+" "+ bevillingsområde+" "+tekstUndersektion)
                selection.insertParagraph('', "After")
            }
          }
        }
      }
    }
    await context.sync()
  });
}


export async function skabelon() {
  return Word.run(async (context) => {
    const valgtDokument = document.getElementById("dokumentDropdown").value;
    const valgtUdvalg = document.getElementById("udvalgDropdown").value;
    
    const responseDokumenttype = await fetch("./assets/dokumenttype.json");
    const dokumenttypeJSON = await responseDokumenttype.json();

    const responseOrganisation = await fetch("./assets/organisation.json");
    const organisationJSON = await responseOrganisation.json();
    
    const dokumentdata=dokumenttypeJSON.filter(obj=>obj.type==valgtDokument);
    const sektioner=dokumentdata[0].sektioner;
    const undersektioner=dokumentdata[0].undersektioner;
    const tabelindhold=dokumentdata[0].tabelindhold;
    console.log(tabelindhold)

    const organisationdata=organisationJSON.filter(obj=>obj.udvalg==valgtUdvalg);
    //console.log(organisationdata)
    const bevillingsområder=[]
    for (var i in organisationdata[0].bevillingsområde) {
      bevillingsområder.push(organisationdata[0].bevillingsområde[i].navn)
    }

    var titel=context.document.body.insertParagraph(valgtDokument, Word.InsertLocation.start)
    titel.font.size=20  



    if (valgtDokument=="Budgetopfølgning") { 

      for (var key in sektioner) {
        if (sektioner.hasOwnProperty(key)) { 
          context.document.body.paragraphs.getLast().select("End")
          var selection = context.document.getSelection()

  
          const sektion = sektioner[key]
  
          if (sektion=="Bevilling") {
            for(var bevillingsområde in bevillingsområder) {
              const tekst=sektion + " " + bevillingsområder[bevillingsområde]
              await context.sync();
              await indsætOverskrift(tekst);
              await indsætUndersektioner(sektion, undersektioner, bevillingsområder[bevillingsområde]);
            } 
          } else {
            await indsætOverskrift(sektion);
          }
        }
      } 

      // Indsætter tabeller for servicerammen
      const contentControls = context.document.contentControls;
      contentControls.load('id');

      await context.sync();

      for(var bevillingsområde in bevillingsområder) {

        const delområder=organisationdata[0].bevillingsområde[bevillingsområde].delområde
        console.log(delområder)

        const ccNavn="Bevilling "+bevillingsområder[bevillingsområde] + " Servicerammen"
        const targetCC=genContentControls.indexOf(ccNavn)
  
        const rækker=delområder.length+1
        const kolonner=tabelindhold[0].kolonnenavneTabel1.length

        var data = [tabelindhold[0].kolonnenavneTabel1]
        for (var delområde in delområder){
          var række=[delområder[delområde]]
          for(var i = 1; i <= kolonner-1; i++) {
            række.push("")
          }
          data.push(række)
          console.log(række)
          console.log(data)
        }

        const table=contentControls.items[targetCC].insertTable(delområder.length+1,tabelindhold[0].kolonnenavneTabel1.length,"start",data);


        table.font.bold=false
        table.font.size=8
        table.autoFitWindow()

        const række1=table.rows.getFirst()
        række1.shadingColor="#DDEBF7"
        række1.verticalAlignment="Center"
        række1.preferredHeight=40
        række1.font.bold=true


        await context.sync();

      }


      
    }
    console.log("nåede hertil")
  });
}

export async function loadContentControls() {
  return Word.run(async (context) =>{
     // Create a proxy object for the content controls collection.
     const contentControls = context.document.contentControls;

     // Queue a command to load the id property for all of the content controls.
     contentControls.load('id');

 
     // Synchronize the document state by executing the queued commands,
     // and return a promise to indicate task completion.
     await context.sync();

     console.log(contentControls.items)

      for (var i in contentControls.items) {
        console.log(contentControls.items[i]._I,)
      }

     if (contentControls.items.length === 0) {
         console.log('No content control found.');
     }
     else {
         // Queue a command to load the properties on the first content control.
         contentControls.items[0].load(  'appearance,' +
                                         'cannotDelete,' +
                                         'cannotEdit,' +
                                         'color,' +
                                         'id,' +
                                         'placeHolderText,' +
                                         'removeWhenEdited,' +
                                         'title,' +
                                         'text,' +
                                         'type,' +
                                         'style,' +
                                         'tag,' +
                                         'font/size,' +
                                         'font/name,' +
                                         'font/color');
 
         // Synchronize the document state by executing the queued commands,
         // and return a promise to indicate task completion.
         await context.sync();
         console.log('Property values of the first content control:' +
             '   ----- appearance: ' + contentControls.items[0].appearance +
             '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
             '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
             '   ----- color: ' + contentControls.items[0].color +
             '   ----- id: ' + contentControls.items[0].id +
             '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
             '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
             '   ----- title: ' + contentControls.items[0].title +
             '   ----- text: ' + contentControls.items[0].text +
             '   ----- type: ' + contentControls.items[0].type +
             '   ----- style: ' + contentControls.items[0].style +
             '   ----- tag: ' + contentControls.items[0].tag +
             '   ----- font size: ' + contentControls.items[0].font.size +
             '   ----- font name: ' + contentControls.items[0].font.name +
             '   ----- font color: ' + contentControls.items[0].font.color);
     }
  })
}

export async function indsætTest() {
  return Word.run(async (context) => {
    
    const contentControls = context.document.contentControls;
    
    contentControls.load('id');

    const targetCC=genContentControls.indexOf('Anlæg')


    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        contentControls.items[targetCC].insertText('Indsat tekst!', 'Replace');
        contentControls.items[targetCC].insertTable(5,5,"Start");
        await context.sync();
    }
  });
}




export async function insertTable() {
  return Word.run(async (context) => {
    // https://www.youtube.com/watch?v=9u6MGqf1J_I

    // Indlæser dokumenttype fra UI
    const dokumenttypeUI = document.getElementById("dokumentDropdown").selectedIndex;
    //console.log(dokumenttypeUI);

    // Indlæser dokumenttype parametre fra json
    const response = await fetch("./assets/dokumenttype.json");
    const dokumenttypeJSON = await response.json();
    console.log(dokumenttypeJSON);

    // Henter kolonneoverskrifter for tabel 1
    const valgtIndex = dokumenttypeUI - 1;
    const dokumenttypeAfgr = dokumenttypeJSON[valgtIndex].tabelindhold;
    //console.log(dokumenttypeAfgr);

    //Udtrækker kolonnenavne for tabel 1
    //for (var key in dokumenttypeAfgr[0].kolonnenavneTabel1) {
    //  context.document.body.insertParagraph(dokumenttypeAfgr[0].kolonnenavneTabel1[key], Word.InsertLocation.end);
    //}
    const antalKolonner = dokumenttypeAfgr[0].kolonnenavneTabel1.length;
    const kolonneNavne = dokumenttypeAfgr[0].kolonnenavneTabel1;
    //console.log(dokumenttypeAfgr[0].kolonnenavneTabel1.length);

    //Udtrækker delområder
    const udvalgUI = document.getElementById("udvalgDropdown").selectedIndex;
    const bevillingsområdeUI = document.getElementById("bevillingsomrDropdown").selectedIndex;

    const responseOrganisation = await fetch("./assets/organisation.json");
    const organisationJSON = await responseOrganisation.json();

    const udvalgIndex = udvalgUI - 1;
    const bevillingsområdeIndex = bevillingsområdeUI - 1;

    const organisationAfgr = organisationJSON[udvalgIndex].bevillingsomr[bevillingsområdeIndex];
    const delområder = organisationAfgr.delområde;
    const antalRækker = organisationAfgr.delområde.length + 1;
    console.log(delområder, antalRækker);

    //const currentYear = new Date(Date.now()).getFullYear();
    //const budgetperiode=[currentYear+1,currentYear+2,currentYear+3,currentYear+4];
    //const overskrift=[""].concat(budgetperiode);

    const data = [kolonneNavne];
    const table = context.document.body.insertTable(antalRækker, antalKolonner, "Start", data);

    const tabelRækker = table.rows;
    tabelRækker.load("items");

    await context.sync();

    for (var i = 1; i <= tabelRækker.items.length; i++) {
      console.log(tabelRækker.items[i].values);
      const rk = (tabelRækker.items[i].values = [[1, 2, 3, 4, 5, 6, 7, 8, 9]]);
      await context.sync();
    }

    await context.sync();
  });
}


export async function addHeader() {
  return Word.run(async (context) => {
    const header1 = document.getElementById("udvalgDropdown").value;
    const header2 = document.getElementById("bevillingsområdeDropdown").value;

    const header = context.document.sections
      .getFirst()
      .getHeader(Word.HeaderFooterType.primary)
      .insertParagraph(header1.concat(" - ", header2), "End");

    header.alignment = "Centered";
    header.font.set({
      bold: false,
      italic: false,
      name: "Calibri",
      color: "black",
      size: 18,
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

    const response = await fetch("./assets/organisation.json");
    const organisation = await response.json();

    // insert a paragraph at the end of the document.
    for (var key in organisation) {
      if (organisation.hasOwnProperty(key)) {
        for (var key2 in organisation[key].bevillingsomr) {
          if (organisation[key].bevillingsomr.hasOwnProperty(key2)) {
            const tekst = organisation[key].udvalg + " - " + organisation[key].bevillingsomr[key2];
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
