import { formaterTabeller } from "./utils/utils.js";


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("skabelon").onclick = skabelon;
    document.getElementById("loadContentControls").onclick = loadElements;
    document.getElementById("rydAlt").onclick = rydAlt;
    document.getElementById("rydAltTools").onclick = rydAlt;
    document.getElementById("rydAltDev").onclick = rydAlt;
    document.getElementById("hentFil").onclick = hentFil;
    document.getElementById("formatterTabeller").onclick = formaterTabeller;

  }
});

export async function loadElements() {
  return Word.run(async (context) => {
    
    var contentControls = context.document.contentControls;
    contentControls.load('items');
    await context.sync();
    console.log("Content controls")
    for (var key in contentControls.items) {
      if (contentControls.items.hasOwnProperty(key)) { 
        console.log(contentControls.items[key]._Ti, ": ",contentControls.items[key]._Te)
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
        console.log(overskriftNiveau,overskrift)
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


export async function formaterTabel(tabel, placering, projekter=0, fodnoteType=0, customFodnote=0) {
  return Word.run(async (context) => {

    tabel.headerRowCount=1
    if (projekter==1) {
      tabel.addRows("end",2,[["I alt ekskl. projekter"],["Projekter"]])
    }
    tabel.addRows("end",1,[["I alt"]])
 
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

export async function hentFil() {
  return Word.run(async (context) => { 
    context.location='ms-word:ofe|u|https://broen.randers.dk/media/60535/tjekliste-it-indkoeb-280621.docx'
    await context.sync();
  })
}
  
// Generer skabelonen
export async function skabelon() {
  return Word.run(async (context) => {
     
    globalThis.genContentControls=[] 
    globalThis.dokumentKommentarer=[]

    const valgtDokument = document.getElementById("dokumentDropdown").value;
    const valgtDokumentDetajle = document.getElementById("dokumentDetaljeDropdown").value;
    var valgtUdvalg = document.getElementById("udvalgDropdown").value;
    
    const responseDokumenttype = await fetch("./assets/dokumenttype.json");
    const dokumenttypeJSON = await responseDokumenttype.json();

    const responseOrganisation = await fetch("./assets/organisation.json");
    const organisationJSON = await responseOrganisation.json();
    
    const dokumentdata=dokumenttypeJSON.filter(obj=>obj.type==valgtDokument);
    const sektioner=dokumentdata[0].sektioner;
    const undersektioner=dokumentdata[0].undersektioner;
    const tabelindhold=dokumentdata[0].tabelindhold;
    const notatDetaljer=dokumentdata[0].notatdetaljer;

    const organisationdata=organisationJSON.filter(obj=>obj.udvalg==valgtUdvalg);
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
    const currentYear=new Date(Date.now()).getFullYear()


    console.log("her: ",afgrænsningsdata[0].undersektioner[0].bevilling[0])
    console.log("org: ",Object.keys(organisationdata[0].bevillingsområde[0]))

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
      console.log(notatDetaljer)
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
      //console.log(dokumentegenskaber)
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

      // Bevillingsområder
      for(var bevillingsområde in bevillingsområder) {
        for (var bevilling in undersektioner[0].bevilling) {      
          var caseVar=undersektioner[0].bevilling[bevilling]
          switch(caseVar) {
            case "Servicerammen":
              // Servicerammen
              const delområder=organisationdata[0].bevillingsområde[bevillingsområde].delområde
            
              var ccNavn="Bevilling "+bevillingsområder[bevillingsområde] +" "+caseVar
              var targetCC=genContentControls.indexOf(ccNavn)

              var rækkerAntal=delområder.length+1
              var kolonnerAntal=tabelindhold[0].kolonnenavneTabelType11.length
              
              // Konstruerer datatabel
              var data = [tabelindhold[0].kolonnenavneTabelType11]
              for (var delområde in delområder){
                var række=[delområder[delområde]]
                for(var i = 1; i <= kolonnerAntal-1; i++) {
                  række.push("")
                }
                data.push(række)
              }

              // Indledende tekst 
              var tabel=contentControls.items[targetCC].insertParagraph(organisationdata[0].bevillingsområde[bevillingsområde].beskrivelse[0],"Start");

              // Datatabel
              var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"End",data);
              formaterTabel(tabel, contentControls.items[targetCC],organisationdata[0].bevillingsområde[bevillingsområde].projekter)
              await context.sync()

              // Tabelbeskrivelse
              tableAltBeskObj(bevillingsområder[bevillingsområde] + " servicerammen", organisationdata[0].bevillingsområde[bevillingsområde].beskrivelse[0])
              await context.sync();

              //// Indsætter undersektioner
              await indsætSektionerICC(ccNavn,delområder,"Heading4");
              //await context.sync();
            ; 
            case "Indkomstoverførsler":
              if (parseInt(bevilling)==1&afgrænsningsdata[0].undersektioner[0].bevilling[bevillingsområde].includes(1)) {
                var ccNavn="Bevilling "+bevillingsområder[bevillingsområde]+" "+caseVar
                var targetCC=genContentControls.indexOf(ccNavn)

                var rækker=[]
                var tempKey=organisationdata[0].bevillingsområde[bevillingsområde].indkomstoverførsler
                for (var i in tempKey) {        
                  rækker.push(tempKey[i])
                }
                var rækkerAntal=rækker.length+1
                var kolonnerAntal=tabelindhold[1].kolonnenavneTabelType12.length
                
                var data = [tabelindhold[1].kolonnenavneTabelType12]
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
        
                tableAltBeskObj(bevillingsområder[bevillingsområde] + " indkomstoverførsler", "Tabellen viser budget, bevillingsansøgninger, forventet forbrug og årets resultat forindkomstoverførsler, "+ bevillingsområder[bevillingsområde])
                await context.sync()

                //// Indsætter undersektioner
                await indsætSektionerICC(ccNavn,rækker,"Heading4"); 
                await context.sync();
        
                //// Sletter tom paragraph før tabel
                var temp=contentControls.items[targetCC].paragraphs.getFirst()
                temp.delete(); 
              }
            ;
            case "Ældreboliger":
              if (parseInt(bevilling)==2&afgrænsningsdata[0].undersektioner[0].bevilling[bevillingsområde].includes(2)) {
                var ccNavn="Bevilling "+bevillingsområder[bevillingsområde]+" "+caseVar
                var targetCC=genContentControls.indexOf(ccNavn)

                var rækker=[]
                var tempKey=organisationdata[0].bevillingsområde[bevillingsområde].ældreboliger
                for (var i in tempKey) {        
                  rækker.push(tempKey[i])
                }
                var rækkerAntal=rækker.length+1
                var kolonnerAntal=tabelindhold[2].kolonnenavneTabelType13.length
                
                var data = [tabelindhold[2].kolonnenavneTabelType13]
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

                tableAltBeskObj(bevillingsområder[bevillingsområde] + " ældreboliger", "Tabellen viser budget, bevillingsansøgninger, forventet forbrug og årets resultat for ældreboliger, "+ bevillingsområder[bevillingsområde])
                await context.sync()
        
                //// Indsætter undersektioner
                await indsætSektionerICC(ccNavn,rækker,"Heading4"); 
                await context.sync();
        
                //// Sletter tom paragraph før tabel
                var temp=contentControls.items[targetCC].paragraphs.getFirst()
                temp.delete(); 
              }
            ;
            case "Brugerfinansieret område":
              if (parseInt(bevilling)==3&afgrænsningsdata[0].undersektioner[0].bevilling[bevillingsområde].includes(3)) {
                var ccNavn="Bevilling "+bevillingsområder[bevillingsområde]+" "+caseVar
                var targetCC=genContentControls.indexOf(ccNavn)

                var rækker=[]
                var tempKey=organisationdata[0].bevillingsområde[bevillingsområde].brugerfinansieret
                for (var i in tempKey) {        
                  rækker.push(tempKey[i])
                }
                var rækkerAntal=rækker.length+1
                var kolonnerAntal=tabelindhold[3].kolonnenavneTabelType14.length
                
                var data = [tabelindhold[3].kolonnenavneTabelType14]
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

                tableAltBeskObj(bevillingsområder[bevillingsområde] + " brugerfinansieret område", "Tabellen viser budget, bevillingsansøgninger, forventet forbrug og årets resultat for det brugerfinansierde område, "+ bevillingsområder[bevillingsområde])
                await context.sync()
        
                //// Indsætter undersektioner
                await indsætSektionerICC(ccNavn,rækker,"Heading4"); 
                await context.sync();
        
                //// Sletter tom paragraph før tabel
                var temp=contentControls.items[targetCC].paragraphs.getFirst()
                temp.delete() 
              }
            ;
            case "Centrale refusionsordninger mv.":
              if (parseInt(bevilling)==4&afgrænsningsdata[0].undersektioner[0].bevilling[bevillingsområde].includes(4)) {
                var ccNavn="Bevilling "+bevillingsområder[bevillingsområde]+" "+caseVar
                var targetCC=genContentControls.indexOf(ccNavn)

                var rækker=[]
                var tempKey=organisationdata[0].bevillingsområde[bevillingsområde].centralerefusionsordninger
                for (var i in tempKey) {        
                  rækker.push(tempKey[i])
                }
                var rækkerAntal=rækker.length+1
                var kolonnerAntal=tabelindhold[4].kolonnenavneTabelType15.length
                
                var data = [tabelindhold[4].kolonnenavneTabelType15]
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
        
                tableAltBeskObj(bevillingsområder[bevillingsområde] + " centrale refusionsordninger mv.", "Tabellen viser budget, bevillingsansøgninger, forventet forbrug og årets resultat for centrale refusionsordninger mv., "+ bevillingsområder[bevillingsområde])
                await context.sync()
                
                //// Indsætter undersektioner
                await indsætSektionerICC(ccNavn,rækker,"Heading4"); 
                await context.sync();
        
                //// Sletter tom paragraph før tabel
                var temp=contentControls.items[targetCC].paragraphs.getFirst()
                temp.delete(); 
              }
            ;
            case "Aktivitetsbestemt medfinansiering":
              if (parseInt(bevilling)==5&afgrænsningsdata[0].undersektioner[0].bevilling[bevillingsområde].includes(5)) {
                var ccNavn="Bevilling "+bevillingsområder[bevillingsområde]+" "+caseVar
                var targetCC=genContentControls.indexOf(ccNavn)

                var rækker=[]
                var tempKey=organisationdata[0].bevillingsområde[bevillingsområde].aktivitetsbestemt
                for (var i in tempKey) {        
                  rækker.push(tempKey[i])
                }
                var rækkerAntal=rækker.length+1
                var kolonnerAntal=tabelindhold[5].kolonnenavneTabelType16.length
                
                var data = [tabelindhold[5].kolonnenavneTabelType16]
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
        
                tableAltBeskObj(bevillingsområder[bevillingsområde] + " aktivitetsbestemt medfinansiering", "Tabellen viser budget, bevillingsansøgninger, forventet forbrug og årets resultat for aktivitetsbestemt medfinansiering mv., "+ bevillingsområder[bevillingsområde])
                await context.sync()
                
                //// Indsætter undersektioner
                await indsætSektionerICC(ccNavn,rækker,"Heading4"); 
                await context.sync();
        
                //// Sletter tom paragraph før tabel
                var temp=contentControls.items[targetCC].paragraphs.getFirst()
                temp.delete(); 
              }
            ;
          }      
        }
      } 

      
      
      // Anlæg
      if (inkluderSektioner[0].includes(2)) {
        var ccNavn="Anlæg"
        var targetCC=genContentControls.indexOf(ccNavn)

        // Hvis "anlæg" null under et givent udvalg i organisation.json, anvendes blot oplysningen fra dokumenttype.json. 
        // Er "anlæg" ikke null anvendes denne som grundlag for rækkerne       

        var rækker=[]

        // Check for undefined 
        if (organisationdata[0].anlæg!=undefined) {
          var tempKey=organisationdata[0].anlæg
          for (var i in tempKey) {        
            rækker.push(organisationdata[0].anlæg[i])
          }
        } else {
          var tempKey=inkluderUndersektionerFlat[0].anlæg[0]
          for (var i in tempKey) {        
            rækker.push(undersektioner[1].anlæg[tempKey[i]])
          }
        }
        
        var rækkerAntal=rækker.length+1
        var kolonnerAntal=tabelindhold[6].kolonnenavneTabelType17.length
        
        var data = [tabelindhold[6].kolonnenavneTabelType17]
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

        tableAltBeskObj(valgtUdvalg + " anlæg", "Tabellen viser budget, bevillingsansøgninger, forventet forbrug og årets resultat for anlæg, "+valgtUdvalg) 
        await context.sync()

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
      var kolonnerAntal=tabelindhold[7].kolonnenavneTabelType2.length
      
      var data = [tabelindhold[7].kolonnenavneTabelType2]
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

      tableAltBeskObj(valgtUdvalg + " bevillingsansøgninger", "Tabellen viser bevillingsansøgninger i budgetåret og overslagsåret for "+valgtUdvalg)
      await context.sync()

      //// Indsætter undersektioner
      await indsætSektionerICC(ccNavn,rækker,"Heading3"); 
      await context.sync();

      //// Sletter tom paragraph før tabel
      var temp=contentControls.items[targetCC].paragraphs.getFirst()
      temp.delete()

      // Custom tabeller
      var customTabeller=afgrænsningsdata[0].customTabeller



      for (var i in customTabeller) {
        var afsnit=context.document.body.paragraphs.load(['text'])
        await context.sync()
        console.log(afsnit)

        console.log(i, customTabeller[i])
        var rækker=customTabeller[i].rækker
        var kolonner=customTabeller[i].kolonner
        var tabelnr=customTabeller[i].tabelnr
        var rækkerAntal=rækker.length+1
        var kolonnerAntal=kolonner.length

        var ccNavn=customTabeller[i].placering
        console.log(ccNavn)
        var targetP=parseInt(await indlæsAfsnit(ccNavn))
        console.log(targetP)
        var data = [kolonner]
        for (var j in rækker){
          var række=[rækker[j]]
          for(var j = 1; j <= kolonnerAntal-1; j++) {
            række.push("")
          }
          data.push(række)
        }

        var nytAfsnit=afsnit.items[targetP].insertParagraph("","After")
        nytAfsnit.styleBuiltIn="Normal"
        await context.sync()

  
        var tabel=nytAfsnit.insertTable(rækkerAntal,kolonnerAntal,"After",data);
        await context.sync()

        var tabeller=context.document.body.tables.load()        
        await context.sync()

        tabeller.items[tabelnr].select("end")
        var placering=context.document.getSelection()
       
        await formaterTabel(tabel,placering,0,2,"") 
       
        tableAltBeskObj(valgtUdvalg + " CT" +i, customTabeller[i].indledendeTekst,customTabeller[i].tabelnr)
        await context.sync()

      } 
      
    }
    
    console.log("nåede hertil")
    formaterTabeller();
  });
}
