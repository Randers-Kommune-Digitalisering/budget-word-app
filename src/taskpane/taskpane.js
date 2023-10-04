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
    document.getElementById("skabelon").onclick = skabelon;
    document.getElementById("insertTable").onclick = insertTable;
    document.getElementById("addHeader").onclick = addHeader;
    document.getElementById("loadContentControls").onclick = loadContentControls;
    document.getElementById("indsætTest").onclick = indsætTest;
    document.getElementById("rydAlt").onclick = rydAlt;
    document.getElementById("indlæsAfsnit").onclick = indlæsAfsnit;
  }
});

export async function indlæsAfsnit(placering) {
  return Word.run(async (context) => {
    var afsnit=context.document.body.paragraphs.load(['text','style'])
    await context.sync()
    var items=afsnit.items
    var overskrift=[]
    var overskriftNiveau=[]
    for(var i in items) {
      if(items[i].style.slice(0,10)=="Overskrift") {
        var nyOverskrift=items[i].text
        var nyOverskriftNiveau=items[i].style.slice(-1)
        if(nyOverskriftNiveau==overskriftNiveau.slice(-1)) {
          overskrift.pop()
          overskriftNiveau.pop()
          overskrift.push(nyOverskrift)
          overskriftNiveau.push(nyOverskriftNiveau)
        }
        if (nyOverskriftNiveau>overskriftNiveau.slice(-1)) {
          overskrift.push(nyOverskrift)
          overskriftNiveau.push(nyOverskriftNiveau)
        }
        if (nyOverskriftNiveau<overskriftNiveau.slice(-1)) {
          while(nyOverskriftNiveau<overskriftNiveau.slice(-1)) {
            overskrift.pop()
            overskriftNiveau.pop()
          }
          overskrift.pop()
          overskriftNiveau.pop()
          overskrift.push(nyOverskrift)
          overskriftNiveau.push(nyOverskriftNiveau)
        }
        //console.log(overskriftNiveau,overskrift.slice(1,overskrift.length).toString().replaceAll(","," "))

        console.log(i ,items[i])
        if (placering==overskrift.slice(1,overskrift.length).toString().replaceAll(","," ")) {
          return i
        }
        
        //return items[i]
        
        //const indsæt=items[i].insertParagraph("Test","After")
        //indsæt.styleBuiltIn="Normal"
      
      }
    }
  })
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

export async function indsætSektion(tekst) {
  return Word.run(async (context) => {
    context.document.body.paragraphs.getLast().select("End")
    var selection = context.document.getSelection()

    
    const overskrift=selection.insertParagraph(tekst);
    overskrift.styleBuiltIn="Heading2"

    await context.sync();
    await indsætConcentControl(tekst)
    
    selection.insertParagraph('', "After");   

  })
}

export async function indsætUndersektionerOld(sektion, undersektioner, ekstraTekst, heading) {
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
              //console.log(tekstUndersektion)
                var underoverskrift=selection.insertParagraph(tekstUndersektion)
                underoverskrift.styleBuiltIn=heading
                context.document.body.paragraphs.getLast().select("End")
                await context.sync();
                await indsætConcentControl(sektion+" "+ ekstraTekst +" "+ tekstUndersektion)
                selection.insertParagraph('', "After")
            }
          }
        }
      }
    }
    await context.sync()
  });
}

export async function indsætUndersektioner(sektion, undersektioner, ekstraTekst, heading) {
  return Word.run(async (context) => {
    context.document.body.paragraphs.getLast().select("End")
    var selection = context.document.getSelection()
    for (var key in undersektioner) {
      const tekstUndersektion=undersektioner[key]
      var underoverskrift=selection.insertParagraph(tekstUndersektion)
      underoverskrift.styleBuiltIn=heading
      context.document.body.paragraphs.getLast().select("End")
      await context.sync();
      await indsætConcentControl(sektion+" "+ ekstraTekst +" "+ tekstUndersektion)
      selection.insertParagraph('', "After")
    }  
    await context.sync()
  });
}

export async function indsætSektionerICC(cc, undersektioner, heading) {
  return Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load('id');

    await context.sync()

    const targetCC=genContentControls.indexOf(cc)
    //const selection=contentControls.items[targetCC].select("End")
    const last=contentControls.items[targetCC]
    const undersektionerRev=undersektioner.slice().reverse()
    for(var undersektion in undersektioner) {
      last.insertParagraph(undersektioner[undersektion],"End")  
      .styleBuiltIn=heading;
      last.insertParagraph('',"End")
      .styleBuiltIn="Normal"
    }

    await context.sync()
  });
}

export async function formaterTabel(tabel, placering, projekter=0, fodnoteType=0,customFodnote=0) {
  return Word.run(async (context) => {
    tabel.font.bold=false
    tabel.font.size=8
    tabel.headerRowCount=1
    if (projekter==1) {
      tabel.addRows("end",2,[["I alt ekskl. projekter"],["Projekter"]])
    }
    tabel.addRows("end",1,[["I alt"]])

    const rækker=tabel.rows
    const række1=rækker.getFirst()
    række1.shadingColor="#DDEBF7"
    række1.verticalAlignment="Center"
    række1.preferredHeight=40
    række1.font.bold=true

    if (customFodnote==0) {
      if (fodnoteType==0) {
        var fodnote=placering.insertText("Note: Minus angiver et mindreforbrug/overskud i Årets forventede resultat og overførsler. Plus angiver et merforbrug/underskud.","End")
        fodnote.font.size=8
        fodnote.font.italic=true
      }
      if (fodnoteType==1) {
        var fodnote=placering.insertText("Note: Minus angiver indtægter, plus angiver udgifter.","End")
        fodnote.font.size=8
        fodnote.font.italic=true
      }
    } else {
      var fodnote=placering.insertText(customFodnote,"End")
      fodnote.font.size=8
      fodnote.font.italic=true
    }
  })
}

export async function skabelon() {
  return Word.run(async (context) => {

    globalThis.genContentControls=[]

    const valgtDokument = document.getElementById("dokumentDropdown").value;
    var valgtUdvalg = document.getElementById("udvalgDropdown").value;
    
    const responseDokumenttype = await fetch("./assets/dokumenttype.json");
    const dokumenttypeJSON = await responseDokumenttype.json();

    const responseOrganisation = await fetch("./assets/organisation.json");
    const organisationJSON = await responseOrganisation.json();
    
    const dokumentdata=dokumenttypeJSON.filter(obj=>obj.type==valgtDokument);
    const sektioner=dokumentdata[0].sektioner;
    const undersektioner=dokumentdata[0].undersektioner;
    const tabelindhold=dokumentdata[0].tabelindhold;

    const organisationdata=organisationJSON.filter(obj=>obj.udvalg==valgtUdvalg);
    //console.log(organisationdata)
    const bevillingsområder=[]
    for (var i in organisationdata[0].bevillingsområde) {
      bevillingsområder.push(organisationdata[0].bevillingsområde[i].navn)
    }

    // Indlæser sektionsafgrænsninger
    const afgrænsningsdata=organisationdata[0].dokumenter.filter(obj=>obj.navn=valgtDokument)
    const inkluderSektioner=[]
    for (var i in afgrænsningsdata[0].sektioner) { 
      inkluderSektioner.push(afgrænsningsdata[0].sektioner[i])
    }

    const inkluderUndersektioner=[]
    for (var i in afgrænsningsdata[0].undersektioner) {
      inkluderUndersektioner.push([afgrænsningsdata[0].undersektioner[i]])
    }
    const inkluderUndersektionerFlat=inkluderUndersektioner.flat(Infinity)
    // console.log(inkluderUndersektionerFlat)

    // Indsætter titel
    var titel=context.document.body.insertParagraph(dokumentdata[0].langtNavn, Word.InsertLocation.start)
    titel.styleBuiltIn="Heading1"

    if (valgtDokument=="Budgetopfølgning") {         
      // Indsætter sektioner og undersektioner
      for (var key in sektioner) {
        if (sektioner.hasOwnProperty(key)) { 
          context.document.body.paragraphs.getLast().select("End")
  
          const sektion = sektioner[key]

          await context.sync(); 
          if (inkluderSektioner[0].includes(parseInt(key))) {
            if (sektion=="Bevilling") {
              for(var bevillingsområde in bevillingsområder) { 
                await indsætSektion(sektion+" "+bevillingsområder[bevillingsområde]);
                await context.sync();              

                //console.log(bevillingsområde)
                // console.log(inkluderUndersektionerFlat[0].bevilling[bevillingsområde])
                
                const inkluderedeUndersektioner=[]
                const inkluderedeUndersektionerKey=inkluderUndersektionerFlat[0].bevilling[bevillingsområde]
                for (var i in inkluderedeUndersektionerKey) {        
                  inkluderedeUndersektioner.push(undersektioner[0].bevilling[inkluderedeUndersektionerKey[i]])
                }
                
                await indsætUndersektioner(sektion, inkluderedeUndersektioner, bevillingsområder[bevillingsområde], "Heading3");
                //await indsætUndersektioner(sektion, undersektioner, bevillingsområder[bevillingsområde], "Heading3");
                await context.sync();
              } 
            } else {
              await indsætSektion(sektion);
            }
          } 
        } 
      } 

      // Indsætter indhold i rammestrukturen
      var contentControls = context.document.contentControls;
      contentControls.load('id');

      await context.sync();

      // Bevillingsområder
      for(var bevillingsområde in bevillingsområder) {
        for (var bevilling in undersektioner[0].bevilling) {      
          var caseVar=undersektioner[0].bevilling[bevilling]
          // console.log("caseVar", caseVar)
          // console.log(caseVar=undersektioner[0].bevilling[bevilling])
          switch(caseVar) {
            case "Servicerammen":
              // Servicerammen
              const delområder=organisationdata[0].bevillingsområde[bevillingsområde].delområde
            
              var ccNavn="Bevilling "+bevillingsområder[bevillingsområde] + " Servicerammen"
              var targetCC=genContentControls.indexOf(ccNavn)

              var rækkerAntal=delområder.length+1
              var kolonnerAntal=tabelindhold[0].kolonnenavneTabelType1.length

              // Konstruerer datatabel
              var data = [tabelindhold[0].kolonnenavneTabelType1]
              for (var delområde in delområder){
                var række=[delområder[delområde]]
                for(var i = 1; i <= kolonnerAntal-1; i++) {
                  række.push("")
                }
                data.push(række)
              }

              var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"Start" ,data);
              await formaterTabel(tabel,contentControls.items[targetCC])
              await context.sync();

              //// Indsætter undersektioner
              await indsætSektionerICC(ccNavn,delområder,"Heading4");
              await context.sync();

              //// Sletter tom paragraph før tabel
              var temp=contentControls.items[targetCC].paragraphs.getFirst()
              temp.delete();
            ;
            case "Brugerfinansieret område":
              if (parseInt(bevilling)==3&afgrænsningsdata[0].undersektioner[0].bevilling[bevillingsområde].includes(3)) {
                var ccNavn="Bevilling "+bevillingsområder[bevillingsområde]+" "+caseVar
                var targetCC=genContentControls.indexOf(ccNavn)

                var rækker=[]
                var tempKey=organisationdata[0].bevillingsområde[0].brugerfinansieret
                for (var i in tempKey) {        
                  rækker.push(tempKey[i])
                }
                var rækkerAntal=rækker.length+1
                var kolonnerAntal=tabelindhold[0].kolonnenavneTabelType1.length
                
                var data = [tabelindhold[0].kolonnenavneTabelType1]
                var række=[]
                for (var i in rækker){
                  var række=[rækker[i]]
                  for(var i = 1; i <= kolonnerAntal-1; i++) {
                    række.push("")
                  }
                  data.push(række)
                }
                
                var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"start",data);
                await formaterTabel(tabel,contentControls.items[targetCC])
                await context.sync();
        
                //// Indsætter undersektioner
                await indsætSektionerICC(ccNavn,rækker,"Heading4"); 
                await context.sync();
        
                //// Sletter tom paragraph før tabel
                var temp=contentControls.items[targetCC].paragraphs.getFirst()
                temp.delete() 
              }
            ;
            case "Centrale refusionsordninger mv.":
            ;
          }
        }      
      }

      // Anlæg
      if (inkluderSektioner[0].includes(2)) {
        var ccNavn="Anlæg"
        var targetCC=genContentControls.indexOf(ccNavn)

        var rækker=[]
        var tempKey=inkluderUndersektionerFlat[0].anlæg[0]
        for (var i in tempKey) {        
        rækker.push(undersektioner[1].anlæg[tempKey[i]])
        }
        var rækkerAntal=rækker.length+1
        var kolonnerAntal=tabelindhold[0].kolonnenavneTabelType1.length
        
        var data = [tabelindhold[0].kolonnenavneTabelType1]
        var række=[]
        for (var i in rækker){
          var række=[rækker[i]]
          for(var i = 1; i <= kolonnerAntal-1; i++) {
            række.push("")
          }
          data.push(række)
        }

        var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"start",data);
        await formaterTabel(tabel,contentControls.items[targetCC])
        await context.sync();

        //// Indsætter undersektioner
        await indsætSektionerICC(ccNavn,rækker,"Heading3"); 
        await context.sync();

        //// Sletter tom paragraph før tabel
        var temp=contentControls.items[targetCC].paragraphs.getFirst()
        temp.delete()
      }

      // Bevillingsansøgninger
      var ccNavn="Bevillingsansøgninger"
      var targetCC=genContentControls.indexOf(ccNavn)

      var rækker=[]
      var tempKey=inkluderUndersektionerFlat[0].bevillingsansøgninger[0]
      for (var i in tempKey) {        
       rækker.push(undersektioner[2].bevillingsansøgninger[tempKey[i]])
      }
      var rækkerAntal=rækker.length+1
      var kolonnerAntal=tabelindhold[1].kolonnenavneTabelType2.length
      
      var data = [tabelindhold[1].kolonnenavneTabelType2]
      for (var i in rækker){
        var række=[rækker[i]]
        for(var i = 1; i <= kolonnerAntal-1; i++) {
          række.push("")
        }
        data.push(række)
      }

      var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"start",data);
      await formaterTabel(tabel,contentControls.items[targetCC],0,1)
      await context.sync();

      //// Indsætter undersektioner
      await indsætSektionerICC(ccNavn,rækker,"Heading3"); 
      await context.sync();

      //// Sletter tom paragraph før tabel
      var temp=contentControls.items[targetCC].paragraphs.getFirst()
      temp.delete()

      // Custom tabeller
      var customTabeller=afgrænsningsdata[0].customTabeller

      var afsnit=context.document.body.paragraphs.load(['text'])
      await context.sync()

      for (var i in customTabeller) {
        var rækker=customTabeller[i].rækker
        var kolonner=customTabeller[i].kolonner
        var tabelnr=customTabeller[i].tabelnr
        var rækkerAntal=rækker.length+1
        var kolonnerAntal=kolonner.length

        var ccNavn=customTabeller[i].placering
        var targetP=parseInt(await indlæsAfsnit(ccNavn))
        var data = [kolonner]
        for (var i in rækker){
          var række=[rækker[i]]
          for(var i = 1; i <= kolonnerAntal-1; i++) {
            række.push("")
          }
          data.push(række)
        }
        //console.log(afsnit.items[targetP], afsnit.items[targetP].text, targetP)
        // var cc=afsnit.items[targetP].select("End")
        // var selection = context.document.getSelection()
        // var cc=selection.insertContentControl("RichText")
        // cc.title="Customtabel"
        const nytAfsnit=afsnit.items[targetP].insertParagraph("","After")
        nytAfsnit.styleBuiltIn="Normal"
        //const nytAfsnit2=nytAfsnit.insertParagraph("","After")
        var tabel=nytAfsnit.insertTable(rækkerAntal,kolonnerAntal,"After",data);
        var tabeller=context.document.body.tables.load()
        //var sidsteTabel=context.document.body.tables.load("id")
        await context.sync()


        // for (var i in tabeller.items) {
        //   tabeller.items[i].select("End")
        //   var selection=context.document.getSelection()
        //   selection.insertText(i,"end")
        // }
        console.log(tabelnr)
        tabeller.items[tabelnr].select("end")
        var placering=context.document.getSelection()
        console.log(placering) 
        //console.log(tabelid)
        //console.log(Math.max(...tabelid))
        //console.log(sidsteTabel.items)
        //sidsteTabel.getLast().select("End")
        //var placering=context.document.getSelection() 


        //var placering=tabel.select("End")
        //afsnit.items[targetP].insertParagraph("","After")
        //var afsnit=context.document.body.paragraphs.load(['text'])
        //var tabel=afsnit.items[targetP+1].insertTable(rækkerAntal,kolonnerAntal,"Before",data);
       
        await formaterTabel(tabel,placering,0,2,"") 
       
        //tabel.insertText("test1","End ")
        await context.sync(); 
  
        //// Indsætter undersektioner
        //await indsætSektionerICC(ccNavn,rækker,"Heading3"); 
        await context.sync();
  
        //// Sletter tom paragraph før tabel
        // var temp=contentControls.items[targetCC].paragraphs.getFirst()
        // temp.delete()
      }
      //console.log(await indlæsAfsnit(ccNavn))
      // var test=await new indlæsAfsnit()
      // context.sync()
      // console.log(test)

      // test.insertParagraph("test","After")
      // console.log("funktion "+await indlæsAfsnit())

    }
    console.log("nåede hertil")
  });
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
    //console.log(dokumenttypeJSON);

    // Henter kolonneoverskrifter for tabel 1
    const valgtIndex = dokumenttypeUI - 1;
    const dokumenttypeAfgr = dokumenttypeJSON[valgtIndex].tabelindhold;
    //console.log(dokumenttypeAfgr);

    //Udtrækker kolonnenavne for tabel 1
    //for (var key in dokumenttypeAfgr[0].kolonnenavneTabelType1) {
    //  context.document.body.insertParagraph(dokumenttypeAfgr[0].kolonnenavneTabelType1[key], Word.InsertLocation.end);
    //}
    const antalKolonner = dokumenttypeAfgr[0].kolonnenavneTabelType1.length;
    const kolonneNavne = dokumenttypeAfgr[0].kolonnenavneTabelType1;
    //console.log(dokumenttypeAfgr[0].kolonnenavneTabelType1.length);

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

    //const currentYear = new Date(Date.now()).getFullYear();
    //const budgetperiode=[currentYear+1,currentYear+2,currentYear+3,currentYear+4];
    //const overskrift=[""].concat(budgetperiode);

    const data = [kolonneNavne];
    const table = context.document.body.insertTable(antalRækker, antalKolonner, "Start", data);

    const tabelRækker = table.rows;
    tabelRækker.load("items");

    await context.sync();

    for (var i = 1; i <= tabelRækker.items.length; i++) {
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
