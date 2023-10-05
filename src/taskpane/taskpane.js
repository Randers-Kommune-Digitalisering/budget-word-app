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
    document.getElementById("rydAlt").onclick = rydAlt;
    document.getElementById("addImage").onclick = addImage;
  }
});

export async function addImage() {
  return Word.run(async (context) => {
    const image=context.document.body.paragraphs
    .getLast()
    .insertParagraph("", "After")
    .insertInlinePictureFromBase64(
      "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"
      ,"Start"
    );
    image
    
  await context.sync();
  })
}


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
