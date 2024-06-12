/* eslint-disable no-prototype-builtins */
/* eslint-disable no-undef */
//import { ContextExclusionPlugin } from "webpack";
import { formaterTabeller, formaterTabellerBB, replaceWordsWithLinks } from "./utils/utils.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("skabelon").onclick = skabelon;
    document.getElementById("loadContentControls").onclick = loadElements;
    document.getElementById("rydAlt").onclick = rydAlt;
    document.getElementById("rydSidehoved").onclick = rydSidehoved;
    document.getElementById("rydAltTools").onclick = rydAlt;
    document.getElementById("rydAltDev").onclick = rydAlt;
    document.getElementById("formaterTabeller").onclick = formaterTabeller;
  }
});

export async function loadElements() {
  return Word.run(async (context) => {
    var contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    console.log("Content controls");
    for (var key in contentControls.items) {
      if (contentControls.items.hasOwnProperty(key)) {
        console.log(contentControls.items[key]._Ti, ": ", contentControls.items[key]._Te);
      }
    }

    var afsnit=context.document.body.paragraphs;
    afsnit.load('items');
    await context.sync();
    console.log("Afsnit")
    for (var key in afsnit.items) {
      if (afsnit.items.hasOwnProperty(key)) { 
        console.log(afsnit.items[key]._Te)
      }
    }

    var tables=context.document.body.tables;
    tables.load('items');
    await context.sync();
    console.log("Tabeller")
    console.log(tables)
    for (var key in tables.items) {
      if (tables.items.hasOwnProperty(key)) { 
        console.table(tables.items[key]._V)
      }
    }
    
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

        if (placering==overskrift.slice(1,overskrift.length).toString().replaceAll(","," ")) {
          return i
        }
      }
    }
  })
}
 

export async function rydSidehoved() { 
  return Word.run(async (context) => {
    // Ryd tekst i header
    var header=context.document.sections.getFirst().getHeader("primary")

    var afsnit=header.paragraphs
    context.load(afsnit, "text");
    await context.sync();
    for (var i = 0; i < afsnit.items.length; i++) {
      afsnit.items[i].delete();
    }

    // Ryd tekst i header
    var footer=context.document.sections.getFirst().getFooter("primary")

    var afsnit=footer.paragraphs
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
  });
}

export async function indsætContentControl(name) {
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
    await indsætContentControl(tekst)
    
    selection.insertParagraph('', "After");   
  })
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
      await indsætContentControl(sektion+" "+ ekstraTekst +" "+ tekstUndersektion)
      selection.insertParagraph('', "After")
    }  
    await context.sync()
  });
}

// Funktion til at indsætte sektioner i contentcontrols
export async function indsætSektionerICC(cc, undersektioner, heading) {
  return Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load('id');

    await context.sync();

    const targetCC=genContentControls.indexOf(cc)
    const last=contentControls.items[targetCC]
    var temp=last.insertParagraph('',"End")
      .styleBuiltIn="Normal";
      for(var undersektion in undersektioner) {
        if (undersektioner.length>1) {
          last.insertParagraph(undersektioner[undersektion],"End")  
          .styleBuiltIn=heading;
        }
          last.insertParagraph('',"End")
          .styleBuiltIn="Normal";
      }
    await context.sync()
  });
}

// Funktion til at gemme et JSON-objekt i dokumentkommentarerne til viderebehandling i VBA. 
export async function tableAltBeskObj(titel,beskrivelse,tabelnr=0) {
  return Word.run(async (context) => {
              
    var temp=context.document.properties.load("comments")
    await context.sync()
    if (temp.comments!="") {
      var tempjson=JSON.parse(temp.comments) 
      var i=tempjson[Object.keys(tempjson)[Object.keys(tempjson).length - 1]].nr  
      i++
    } else {
      var i=1
    }
    
    if (tabelnr==0) {
      dokumentKommentarer.push({nr:i,titel:titel,beskrivelse:beskrivelse})
    } else {
      dokumentKommentarer.splice(tabelnr-1,0,{nr:i,titel:titel,beskrivelse:beskrivelse})
    }

    // Renummerer tabeller (custom tabeller kan sættes ind over alt)
    for (var j=0; j<dokumentKommentarer.length; j++) {
      dokumentKommentarer[j].nr=j+1
    }

    context.document.properties.set({comments:JSON.stringify(dokumentKommentarer)})          
    await context.sync()  
  })
}


export async function tabelAddOns(tabel, placering, projekter=0, fodnote=0) {
  return Word.run(async (context) => {

    tabel.headerRowCount=1
    if (projekter==1) {
      tabel.addRows("end",2,[["I alt ekskl. projekter"],["Projekter"]])
    }
    tabel.addRows("end",1,[["I alt"]])

    if (fodnote!=0) {
      var indsatFodnote=placering.insertText(fodnote,"End")
      indsatFodnote.font.size=8
      indsatFodnote.font.italic=true
    }
  })
}

export async function fetchAssets(adr) {
  return Word.run(async (context) => {
    var response = await fetch(adr, { cache: "reload" });
    return response.json();
  });
};


  
// Generer skabelonen
export async function skabelon() {
  return Word.run(async (context) => {

    /* 
    Dette virker ikke. Se OneNote/ChatGPT for alternativ løsning.
    https://chatgpt.com/share/8f0b37b7-d4f4-4e8b-adb9-992f549b29a3
    var response = await fetch("file:////randers.dk/dfs/Budget/Karl Fritjof Krassel/test.csv", { cache: "reload" });
    console.log(response)
    console.table(response.text())
    */ 

    globalThis.genContentControls=[] 
    globalThis.dokumentKommentarer=[]

    const valgtDokument = document.getElementById("dokumentDropdown").value;
    const valgtDokumentDetajle = document.getElementById("dokumentDetaljeDropdown").value;
    const valgtUdvalg = document.getElementById("udvalgDropdown").value;
    const valgtBevilling = document.getElementById("bevillingsområdeDropdown").value;
    
    const responseDokumenttype = await fetch("./assets/dokumenttype.json", { cache: "reload" });
    const dokumenttypeJSON = await responseDokumenttype.json();

    //const responseOrganisation = await fetch("./assets/organisation.json", { cache: "reload" }); 
    //const organisationJSON = await responseOrganisation.json();
    
    // console.log("fetchAssets(\"./assets/organisation.json\"): ", await fetchAssets("./assets/organisation.json"))
    // console.log("OrganisationJSON: ", organisationJSON)

    // const dokumentdata=organisationJSON.filter(obj=>obj.type==valgtDokument); 


    // var responseOrganisation_bbd1 = await fetch("./assets/organisation_bbd1.json");
    // var organisationJSON_bbd1 = await responseOrganisation_bbd1.json();
    // console.log("organisationJSON_bbd1: ",organisationJSON_bbd1) 
    
    const dokumentdata=dokumenttypeJSON.filter(obj=>obj.type==valgtDokument);
    const sektioner=dokumentdata[0].sektioner;
    const undersektioner=dokumentdata[0].undersektioner;
    const tabelindhold=dokumentdata[0].tabelindhold;
    const notatDetaljer=dokumentdata[0].notatdetaljer; 
    const langtNavn=dokumentdata[0].langtNavn; 

    var organisationJSON=await fetchAssets("./assets/organisation.json")
    var organisationdata = organisationJSON.filter(obj=>obj.udvalg==valgtUdvalg);
    //console.log("organisationdata: ",organisationdata)

    /* Udlæser bevillingsområder fra første dokumenttype - ændrer sig ikke på tværs af typer*/
    const bevillingsområder=[] 
    for (var i in organisationdata[0].dokumenter[0].bevillingsområde) {
      bevillingsområder.push(organisationdata[0].dokumenter[0].bevillingsområde[i].navn)
    }

    // Indlæser sektionsafgrænsninger
    var afgrænsningsdata=organisationdata[0].dokumenter.filter(obj=>obj.navn=valgtDokument)
    console.log("afgrænsningsdata: ",afgrænsningsdata) 
    const inkluderSektioner=[]
    for (var i in afgrænsningsdata[0].sektioner) {  
      inkluderSektioner.push(afgrænsningsdata[0].sektioner[i])
    }

    const inkluderUndersektioner=[]
    for (var i in afgrænsningsdata[0].undersektioner) {
      inkluderUndersektioner.push([afgrænsningsdata[0].undersektioner[i]])
    }
    const inkluderUndersektionerFlat=inkluderUndersektioner.flat(Infinity)
    const currentYear=new Date(Date.now()).getFullYear()
    const budgetperiodeÅr1=currentYear+1
    const budgetperiodeÅr2=currentYear+2
    const budgetperiodeÅr3=currentYear+3
    const budgetperiodeÅr4=currentYear+4
    const budgetperiode= budgetperiodeÅr1+"-"+ budgetperiodeÅr4

    if (valgtDokument=="Budgetopfølgning") {        

      // Indsætter notattitel
      const notatTitel=context.document.body.insertParagraph("Budgetopfølgning pr. "+valgtDokumentDetajle+" "+currentYear, Word.InsertLocation.start)
      notatTitel.style="Brev/notat KORT (O1)";
         
      // Indsætter notatdetaljer
      await indsætContentControl("Notatdetaljer")
      context.document.body.paragraphs.getLast().select("End")
      var selection = context.document.getSelection()
      selection.insertParagraph('', "After");   

      const cc = context.document.contentControls;
      cc.load("items")
      await context.sync()
      for (var i in notatDetaljer) {
        var tekst=cc.items[0].insertParagraph(notatDetaljer[i],"End")
        tekst.set({
          lineUnitAfter:0,
          lineUnitBefore:0,
          spaceAfter:0,
          spaceBefore:0
        })         
        tekst.font.set({
          bold: true          
        })
        // Brugeren skriver videre i ikke-fed tekst.  
        var tekstUdfyld=tekst.insertText("	","End")
        tekstUdfyld.font.set({
          bold: false          
        })
      }

      // Sletter tom første paragraph
      var temp=cc.items[0].paragraphs.getFirst()
      temp.delete()

      // Indsæter titel
      var titel=context.document.body.insertParagraph(valgtUdvalg+" – "+dokumentdata[0].langtNavn.toLowerCase()+" pr. "+valgtDokumentDetajle+" "+currentYear, Word.InsertLocation.end)
      titel.styleBuiltIn="Heading1"

      // Indsætter dokumenttitel
      var dokumentegenskaber=context.document.properties.load("title")
      await context.sync()
      context.document.properties.set({title:valgtUdvalg+" – "+dokumentdata[0].langtNavn.toLowerCase()+" pr. "+valgtDokumentDetajle+" "+currentYear})
      
      await context.sync();
      
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
                
                const inkluderedeUndersektioner=[]
                const inkluderedeUndersektionerKey=inkluderUndersektionerFlat[0].bevilling[bevillingsområde]
                for (var i in inkluderedeUndersektionerKey) {        
                  inkluderedeUndersektioner.push(undersektioner[0].bevilling[inkluderedeUndersektionerKey[i]])
                }
                
                await indsætUndersektioner(sektion, inkluderedeUndersektioner, bevillingsområder[bevillingsområde], "Heading3");
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

      // Service
      // Bevillingsområder
      for(var bevillingsområde in bevillingsområder) {
        // Tabeller for hvert bevillingsområde
        const tabeller=afgrænsningsdata[0].bevillingsområde[bevillingsområde].tabeller
        for (var tabel in tabeller) {    
          const ccNavn="Bevilling "+bevillingsområder[bevillingsområde] +" "+tabeller[tabel].navn
          const targetCC=genContentControls.indexOf(ccNavn)
    
          // Indledende tekst 
          const tekst=contentControls.items[targetCC].insertParagraph(tabeller[tabel].beskrivelse,"Start");

          // Datatabel
          var rækker=tabeller[tabel].rækker
          var rækkerAntal=tabeller[tabel].rækker.length+1
          var kolonner=tabelindhold[tabeller[tabel].typeKolonner].overskrifter
          var kolonnerAntal=kolonner.length
          var projekter=tabeller[tabel].projekter
          var fodnote=tabeller[tabel].note

          var data = [kolonner]
          for (var j in rækker){
            var række=[rækker[j]]
            for(var i = 1; i <= kolonnerAntal-1; i++) {
              række.push("")
            }
            data.push(række)
          }
          
         
          var indsatTabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"End",data);
          tabelAddOns(indsatTabel, contentControls.items[targetCC], projekter, fodnote)
    
          // Tabelbeskrivelse i dokumentegenskaber (til VBA-script)
          tableAltBeskObj(bevillingsområder[bevillingsområde] + tabeller[tabel].navn, tabeller[tabel].beskrivelse)
          await context.sync();

          //// Indsætter undersektioner
          await indsætSektionerICC(ccNavn,rækker,"Heading4");
          //await context.sync();

          await context.sync()  
        }
      } 
      
      // Anlæg
      if (inkluderSektioner[0].includes(2)) {
        var ccNavn="Anlæg"
        var targetCC=genContentControls.indexOf(ccNavn)
              
        var anlæg=afgrænsningsdata[0].anlæg[0]

        const tekst=contentControls.items[targetCC].insertParagraph(anlæg.beskrivelse,"Start");

        // Hvis "anlæg" null under et givent udvalg i organisation.json, anvendes blot oplysningen fra dokumenttype.json. 
        // Er "anlæg" ikke null anvendes denne som grundlag for rækkerne       
        var rækker=[] 
        if (anlæg!=undefined) {
          var rækkerTilBrug=anlæg.rækker
        } else {
          var rækkerTilBrug=inkluderUndersektionerFlat[0].anlæg[0]
        }
        for (var række in rækkerTilBrug) {        
          rækker.push(rækkerTilBrug[række])
        }
        var rækkerAntal=rækker.length+1 
        var kolonner=tabelindhold[anlæg.typeKolonner].overskrifter 
        var kolonnerAntal=kolonner.length
        var fodnote=anlæg.note


        var data = [kolonner]
        for (var j in rækker){
          var række=[rækker[j]]
          for(var i = 1; i <= kolonnerAntal-1; i++) {
            række.push("")
          }
          data.push(række)
        }

        var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"End",data);
        await tabelAddOns(tabel,contentControls.items[targetCC],0,fodnote)
        await context.sync();

        tableAltBeskObj(valgtUdvalg + " anlæg", anlæg.beskrivelse) 
        await context.sync()

        //// Indsætter undersektioner
        await indsætSektionerICC(ccNavn,rækker,"Heading3"); 
        await context.sync(); 

      }
      
      // Bevillingsansøgninger
      var ccNavn="Bevillingsansøgninger"
      var targetCC=genContentControls.indexOf(ccNavn)

      var bevillingsansøgninger=afgrænsningsdata[0].bevillingsansøgninger[0]

      var rækker=[] 
      var rækkerTilBrug=bevillingsansøgninger.rækker
      for (var række in rækkerTilBrug) {        
        rækker.push(rækkerTilBrug[række])
      }
      var rækkerAntal=rækker.length+1
      var kolonner=tabelindhold[bevillingsansøgninger.typeKolonner].overskrifter 
      var kolonnerAntal=kolonner.length
      var fodnote=bevillingsansøgninger.note
      
      var data = [kolonner]
      for (var j in rækker){
        var række=[rækker[j]]
        for(var i = 1; i <= kolonnerAntal-1; i++) {
          række.push("")
        }
        data.push(række) 
      } 

      var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"start",data);
      tabelAddOns(tabel,contentControls.items[targetCC],0,fodnote)

      tableAltBeskObj(valgtUdvalg + " bevillingsansøgninger", bevillingsansøgninger.beskrivelse)
      await context.sync()

      // Indsætter undersektioner
      await indsætSektionerICC(ccNavn,rækker,"Heading3"); 
      await context.sync();

      // Sletter tom paragraph før tabel
      var temp=contentControls.items[targetCC].paragraphs.getFirst()
      temp.delete()

      // Custom tabeller
      var customTabeller=afgrænsningsdata[0].customTabeller

      for (var i in customTabeller) {
        var afsnit=context.document.body.paragraphs.load(['text'])
        await context.sync()

        var rækker=customTabeller[i].rækker
        var kolonner=customTabeller[i].kolonner
        var tabelnr=customTabeller[i].tabelnr
        var rækkerAntal=rækker.length+1
        var kolonnerAntal=kolonner.length
        var fodnote=customTabeller[i].note
        var placeringOmkringAfsnit = customTabeller[i].placeringOmkringAfsnit;

        var ccNavn=customTabeller[i].placering
        var targetP=parseInt(await indlæsAfsnit(ccNavn))
        var data = [kolonner]
        for (var j in rækker){
          var række=[rækker[j]]
          for(var j = 1; j <= kolonnerAntal-1; j++) {
            række.push("")
          }
          data.push(række)
        }

        var nytAfsnit = afsnit.items[targetP].insertParagraph("", placeringOmkringAfsnit);
        nytAfsnit.styleBuiltIn="Normal"
        await context.sync()

        var tabel=nytAfsnit.insertTable(rækkerAntal,kolonnerAntal,"Before",data);

        // Fodnoten indsættes selvstændigt for CS-tabeller, da den ellers vil indsættes formert
        var indsatFodnote=nytAfsnit.insertParagraph(fodnote,"Before")
        indsatFodnote.font.size=8
        indsatFodnote.font.italic=true

        await context.sync()

        var tabeller=context.document.body.tables.load()        
        await context.sync()

        tabeller.items[tabelnr].select("start")
        var placering=context.document.getSelection() 
        await context.sync()

        await tabelAddOns(tabel,placering,0,0) 
        await context.sync()

        tableAltBeskObj(valgtUdvalg + " CT" +i, customTabeller[i].indledendeTekst,customTabeller[i].tabelnr)
        await context.sync()

      } 
      formaterTabeller();
    }
    
    if (valgtDokument=="Budgetbemærkninger del 1") {
      
      //  Fetcher organisationsdata igen
      var organisation=await fetchAssets("./assets/organisation.json")
      console.log("organisation: ",organisation)
      var inputdata=organisation.filter(obj=>obj.udvalg==valgtUdvalg)
      inputdata=inputdata[0].dokumenter.filter(obj => obj.navn==valgtDokument)
      inputdata=inputdata[0].bevillingsområde.filter(obj => obj.navn==valgtBevilling)
      console.log("inputdata: ",inputdata)

      // tabelindhold
      var tabeller=inputdata[0].tabeller
      var faktaoverskrift=tabeller[0].faktaOverskrift
      var fakta=tabeller[0].fakta
      var politikker=tabeller[0].politikker
      console.log("tabeller: ",tabeller)  
      
      // Indsætter dokumenttitel
      var dokumentegenskaber=context.document.properties.load("title")
      await context.sync()
      context.document.properties.set({title:valgtUdvalg+" - "+valgtBevilling+" – "+langtNavn.toLowerCase()+" - "+budgetperiode})
      await context.sync();

      // Sidehoved
      // Rydder sidehoved i startskabelonen
      rydSidehoved()

      var header=context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary)
        .insertParagraph(valgtUdvalg+" - "+valgtBevilling,"Start")
      header.font.size=18;
      header.alignment="Centered";
      var header2=context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary)
        .insertParagraph(langtNavn,"End")
      header2.font.size=18;
      header2.alignment="Centered";

      // Indsætter sektioner og undersektioner
      for (var sektion in sektioner) {
        context.document.body.paragraphs.getLast().select("End")

        // Sektioner og undersektioner
        await indsætSektion(sektioner[sektion]);
        await context.sync();           
                
        await indsætUndersektioner(sektioner[sektion], undersektioner[sektion],"-", "Heading3");
        await context.sync();   
      } 
      
      // Indsætter indhold i rammestrukturen
      var contentControls = context.document.contentControls;
      contentControls.load('id');
      await context.sync();

      // Indsætter tabel til fakta og politikker
      rækker=Math.max(fakta.length,politikker.length+2)
      console.log(rækker)
      var data=[]
      data.push([faktaoverskrift,"","","Politikker"])
      
      for (var række=0; række<rækker; række++) {
          data.push([fakta[række],"","",politikker[række]])
      }
      data[rækker-1][3]="Se Randers Kommunes politikker på"
      data[rækker][3]="www.randers.dk/demokrati/politikker/"

      var ccNavn="1. Beskrivelse af området"
      var targetCC=genContentControls.indexOf(ccNavn)
      var indsatTabel=contentControls.items[targetCC].insertTable(rækker+1,4,"End",data);

      tableAltBeskObj("Fakta og politikker", "Fakta og politikker")

      /*
      var cell=indsatTabel.getCell(rækker,3)
      var cellBody=cell.body
      var paragraph = cellBody.paragraphs.getFirst();
      var range = paragraph.getTextRanges(["Randers.dk/politikker"], false).getFirst();

      range.set({hyperlink:'https://github.com'})
      // range.hyperlink="https://randers.dk/politikker"
      //console.log(range)

      */

      // Styler tabel - skal flyttes til utils
      indsatTabel.headerRowCount = 1
      indsatTabel.font.bold=false
      indsatTabel.font.size=11
      indsatTabel.font.name="Calibri"
      indsatTabel.font.color="#000000"

      // Loop over alle rækker
      var rækker=indsatTabel.rows
      rækker.load('items')
      await context.sync()
      for (var i=0; i<rækker.items.length; i++) {
        rækker.items[i].verticalAlignment="Top"
        
        if (i==0) {
          rækker.items[i].font.bold=true
        }

        // Loop over celler
        var celler=rækker.items[i].cells
        celler.load('items')
        await context.sync()
        for (var k=0; k<celler.items.length; k++) {
          
          // Styler kolonne 0, 1 og 3
          if (k!=2) {
            celler.items[k].shadingColor="#DDEBF7"
          }
          // Sætter padding
          celler.items[k].setCellPadding("Top",1)
          celler.items[k].setCellPadding("Bottom",1)

          // Højrestiller kolonne 1
          if (k==1) {
            celler.items[k].horizontalAlignment="Right" 
          } 
          // Indstiller bredden
          // 28.35 points pr. centimer, 17 cm sidebredde. = 491,95 points
          if (k==0) {
            celler.items[k].columnWidth=180  
          }
          if (k==1) {
            celler.items[k].columnWidth=40  
          }
          if (k==2) {
            celler.items[k].columnWidth=42  
          }
          if (k==3) {
            celler.items[k].columnWidth=220  
          }
          if (i==rækker.items.length-1 & k==3) { 
            celler.items[k].verticalAlignment="Top"
            // VIRKER IKKE... 
            /*
            const cell=celler.items[k]
            cell.insertHtml(
              `<a href="Randers.dk/politikker">Randers.dk/politikker</a>`,
              Word.InsertLocation.replace
            );
            */
          }
        }
      }
      

      // Sletter sidste tomme afsnit - virker ikke
      /*
      const afsnit=contentControls.items[targetCC].paragraphs
      afsnit.load('items')
      await context.sync()  
      for (var i=0; i<afsnit.items.length; i++) {
        console.log(afsnit.items[i].text)
      }
      console.log("sidste afsnit",afsnit.items[afsnit.items.length-1])
      
      afsnit.items[afsnit.items.length-1].delete()
      await context.sync()  
      */
      
      // Fjerner alle rammer
      var borderLocation = Word.BorderLocation.all;
      var border = indsatTabel.getBorder(borderLocation);
      border.set({type:'none'})
      await context.sync();

      // Indsætter tabel vedr. drift
      var ccNavn="2. Hovedtal - 2.1 Drift"
      var targetCC=genContentControls.indexOf(ccNavn)

      // Indledende tekst 
      const parse = require('json-templates');
      const templateBeskrivelse = parse(tabeller[1].beskrivelse);
      console.log(templateBeskrivelse({ fra: budgetperiodeÅr1, til: budgetperiodeÅr4 })); 

      const indsatTabelbeskrivelse=contentControls.items[targetCC].insertParagraph(templateBeskrivelse({ fra: budgetperiodeÅr1, til: budgetperiodeÅr4 }),"Start");
      await context.sync();

      var rækkerServicerammen=tabeller[1].rækkerServicerammen
      var rækkerUdenForServicerammen=tabeller[1].rækkerUdenForServicerammen
      var k1r1=tabeller[1].k1r1
      var tabelnr=tabeller[1].nr
      var rækkerServicerammenAntal=rækkerServicerammen.length
      var rækkerUdenForServicerammenAntal=rækkerUdenForServicerammen.length
      var kolonnerAntal=5
      var fodnote=tabeller[1].note

      var data = [[k1r1, budgetperiodeÅr1, budgetperiodeÅr2, budgetperiodeÅr3, budgetperiodeÅr4]] 
      
      // Servicerammen
      data.push(["Servicerammen","","","",""])
      for (var j in rækkerServicerammen){
        var række=[rækkerServicerammen[j]]
        for(var i = 1; i <= kolonnerAntal-1; i++) {
          række.push("")
        }
        data.push(række)
      }
      // Uden for servicerammen
      data.push(["Uden for servicerammen","","","",""])
      for (var j in rækkerUdenForServicerammen){
        var række=[rækkerUdenForServicerammen[j]]
        for(var j = 1; j <= kolonnerAntal-1; j++) {
          række.push("")
        }
        data.push(række)
      }
      
      var indsatTabel=contentControls.items[targetCC].insertTable(rækkerServicerammenAntal+rækkerUdenForServicerammenAntal+3,kolonnerAntal,"End",data);
      //console.table(data)
      await context.sync()
      tabelAddOns(indsatTabel,contentControls.items[targetCC],0,fodnote)
      await context.sync()
      tableAltBeskObj(tabeller[1].navn, templateBeskrivelse({ fra: budgetperiodeÅr1, til: budgetperiodeÅr4 }))
      await context.sync()
      contentControls.items[targetCC].insertParagraph('2.1.1 Servicerammen',"End").styleBuiltIn="Heading3"
      contentControls.items[targetCC].insertParagraph('',"End").styleBuiltIn="Normal"
      contentControls.items[targetCC].insertParagraph('2.1.2 Uden for servicerammen',"End").styleBuiltIn="Heading3"
      contentControls.items[targetCC].insertParagraph('',"End").styleBuiltIn="Normal"

      formaterTabellerBB("tabel-1")

      // Indsætter tabel vedr. anlæg
      var ccNavn="2. Hovedtal - 2.2 Anlæg"
      var targetCC=genContentControls.indexOf(ccNavn)

      // Indledende tekst 
      const parse2 = require('json-templates');
      const templateBeskrivelse2 = parse2(tabeller[2].beskrivelse);
      var indsatTabelbeskrivelse2=contentControls.items[targetCC].insertParagraph(templateBeskrivelse2({ fra: budgetperiodeÅr1, til: budgetperiodeÅr4 }),"Start");
      await context.sync();

      var rækker=tabeller[2].rækker
      var rækkerAntal=tabeller[2].rækker.length
      var k1r1=tabeller[2].k1r1
      var tabelnr=tabeller[2].nr
      var kolonnerAntal=5
      const parse3 = require('json-templates');
      const templateBeskrivelse3 = parse3(tabeller[2].note);

      var data = [[k1r1, budgetperiodeÅr1, budgetperiodeÅr2, budgetperiodeÅr3, budgetperiodeÅr4]] 
      
      for (var j in rækker){
        var række=[rækker[j]]
        for(var i = 1; i <= kolonnerAntal-1; i++) {
          række.push("")
        }
        data.push(række)
      }
      console.log(data) 
      
      var indsatTabel=contentControls.items[targetCC].insertTable(rækkerAntal+1,kolonnerAntal,"End",data);
      await context.sync()
      tabelAddOns(indsatTabel,contentControls.items[targetCC],0,templateBeskrivelse3({ fra: budgetperiodeÅr1, til: budgetperiodeÅr4 }))
      await context.sync()
      tableAltBeskObj(tabeller[2].navn, templateBeskrivelse2({ fra: budgetperiodeÅr1, til: budgetperiodeÅr4 }))
      await context.sync()

      formaterTabellerBB("tabel-2")
      /*


      // var indsatTabel=contentControls.items[targetCC].insertTable(rækker+1,4,"End",data);

        // Datatabel
        var rækker=tabeller[tabel].rækker
        var rækkerAntal=tabeller[tabel].rækker.length+1
        var kolonner=tabelindhold[tabeller[tabel].typeKolonner].overskrifter
        var kolonnerAntal=kolonner.length
        var projekter=tabeller[tabel].projekter
        var fodnote=tabeller[tabel].note

        var data = [kolonner]
        for (var j in rækker){
          var række=[rækker[j]]
          for(var i = 1; i <= kolonnerAntal-1; i++) {
            række.push("")
          }
          data.push(række)
        }
        
        
        var indsatTabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"End",data);
        tabelAddOns(indsatTabel, contentControls.items[targetCC], projekter, fodnote)
  


      var contentControls = context.document.contentControls;
      contentControls.load('items');
      await context.sync();
      */
      
  
    } 
     
    // replaceWordsWithLinks()
 
    console.log("nåede hertil")

  });
}
