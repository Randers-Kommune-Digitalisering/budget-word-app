switch(caseVar) {
    case "Servicerammen":
      // Servicerammen
      const delområder=organisationdata[0].bevillingsområde[bevillingsområde].delområde
    
      var ccNavn="Bevilling "+bevillingsområder[bevillingsområde] +" "+caseVar
      var targetCC=genContentControls.indexOf(ccNavn)

      // Indledende tekst 
      var tabel=contentControls.items[targetCC].insertParagraph(organisationdata[0].bevillingsområde[bevillingsområde].beskrivelse[0],"Start");

      // Datatabel
      // Konstruerer datatabel
      var rækkerAntal=delområder.length+1
      var kolonnerAntal=tabelindhold[0].kolonnenavneTabelType11.length
      
      var data = [tabelindhold[0].kolonnenavneTabelType11]
      for (var delområde in delområder){
        var række=[delområder[delområde]]
        for(var i = 1; i <= kolonnerAntal-1; i++) {
          række.push("")
        }
        data.push(række)
      }

      // Indsætter datatabel
      var tabel=contentControls.items[targetCC].insertTable(rækkerAntal,kolonnerAntal,"End",data);
      tabelAddOns(tabel, contentControls.items[targetCC],organisationdata[0].bevillingsområde[bevillingsområde].projekter)
      await context.sync()

      // Tabelbeskrivelse i dokumentegenskaber (til VBA-script)
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
        await tabelAddOns(tabel,contentControls.items[targetCC])
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
        await tabelAddOns(tabel,contentControls.items[targetCC])
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
        await tabelAddOns(tabel,contentControls.items[targetCC])
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
        await tabelAddOns(tabel,contentControls.items[targetCC])
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
        await tabelAddOns(tabel,contentControls.items[targetCC])
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