export function sumArrays(...arrays) {
  const n = arrays.reduce((max, xs) => Math.max(max, xs.length), 0);
  const result = Array.from({ length: n });
  return result.map((_, i) => arrays.map(xs => xs[i] || 0).reduce((sum, x) => sum + x, 0));
}

// Formatter tabeller
export async function formaterTabeller(){
    return Word.run(async (context) => {
        
      // Loop over alle tabeller
      var tables=context.document.body.tables
      tables.load('items')
      await context.sync()
      if (tables.items.length > 0) {
        for (var j=0;j<tables.items.length;j++) {
          var table = tables.items[j];
          table.headerRowCount = 1
          table.font.bold=false
          table.font.size=8
  
          // Fjerner alle rammer
          var borderLocation = Word.BorderLocation.all;
          var border = table.getBorder(borderLocation);
          border.set({type:'none'})
  
          // Tilføjer horisontale streger
          var borderLocation = Word.BorderLocation.insideHorizontal
          var border = table.getBorder(borderLocation);
          border.set({color:"#D9D9D9",width:1, type:'Single'})
  
          // Loop over alle rækker
          var rækker=table.rows
          rækker.load('items')
          await context.sync()
          for (var i=0; i<rækker.items.length; i++) {
            rækker.items[i].verticalAlignment="Center"
            
            // Styler førster og sidste række
            if (i==0|i==rækker.items.length-1) {
              var borderLocation = Word.BorderLocation.top;
              var border = rækker.items[i].getBorder(borderLocation);
              border.set({color:"#808080",width:1, type:'Single'})
              var borderLocation = Word.BorderLocation.bottom;
              var border = rækker.items[i].getBorder(borderLocation);
              border.set({color:"#808080",width:1, type:'Single'})
              rækker.items[i].shadingColor="#DDEBF7"
              rækker.items[i].font.bold=true
              rækker.items[i].font.name="Calibri"
            }
  
            // Loop over celler
            var celler=rækker.items[i].cells
            celler.load('items')
            await context.sync()
            for (var k=0; k<celler.items.length; k++) {
              // Sætter padding
              celler.items[k].setCellPadding("Top",3)
              celler.items[k].setCellPadding("Bottom",3)
  
              
              // Højrestiller kolonneoverskrifter til højre hvis kun tal
              if (i==0) {
                const årstal=/^\d{4}$/;
                if (årstal.test(celler.items[k].value)==true) {
                  celler.items[k].horizontalAlignment="Right"
                }
              }
              // Højrestiller kolonne > 1 og række > 1
              if (i>0&k>0) {
                celler.items[k].horizontalAlignment="Right" 
              } 
              if (k==0) {
                celler.items[k].columnWidth=240   // Virker vist ikke efter hensigten...
              }
              if (celler.items[k].value.slice(0,5)=="I alt") {
                console.log(celler.items[k].value.slice(0,5))
                rækker.items[i].shadingColor="#DDEBF7"
                rækker.items[i].font.bold=true
                celler.items[k].setCellPadding("Top",3)
                celler.items[k].setCellPadding("Bottom",3)
                var borderLocation = Word.BorderLocation.top;
                var border = rækker.items[i].getBorder(borderLocation);
                border.set({color:"#808080",width:1, type:'Single'})
                var borderLocation = Word.BorderLocation.bottom;
                var border = rækker.items[i].getBorder(borderLocation);
                border.set({color:"#808080",width:1, type:'Single'})
              }
  
            }
          }
        }
        await context.sync()
        }
      await context.sync()
    })
  }

// Formater tabeller til budgetbemærninger
export async function formaterTabellerBB(tabel){
  return Word.run(async (context) => { 
    console.log(tabel)

    const tables = context.document.body.tables;
    tables.load("items");
    
    await context.sync();

    // tilføjer custom tag til hver tabel
    for (let i = 0; i < tables.items.length; i++) {
      tables.items[i].tag = `tabel-${i}`;
    }

    // Afgrænser til relevant tabel
    const table = tables.items.find(table => table.tag === tabel);
    await context.sync();

    table.headerRowCount = 1
    table.font.bold=false
    table.font.size=8

    // Fjerner alle rammer
    var borderLocation = Word.BorderLocation.all;
    var border = table.getBorder(borderLocation);
    border.set({type:'none'})

    // Tilføjer horisontale streger
    var borderLocation = Word.BorderLocation.insideHorizontal
    var border = table.getBorder(borderLocation);
    border.set({color:"#D9D9D9",width:1, type:'Single'})

    // Loop over alle rækker
    var rækker=table.rows
    rækker.load('items')
    await context.sync()
    for (var i=0; i<rækker.items.length; i++) {
      rækker.items[i].verticalAlignment="Center"
      
      // Styler førster og sidste række
      if (i==0|i==rækker.items.length-1) {
        var borderLocation = Word.BorderLocation.top;
        var border = rækker.items[i].getBorder(borderLocation);
        border.set({color:"#808080",width:1, type:'Single'})
        var borderLocation = Word.BorderLocation.bottom;
        var border = rækker.items[i].getBorder(borderLocation);
        border.set({color:"#808080",width:1, type:'Single'})
        rækker.items[i].shadingColor="#DDEBF7"
        rækker.items[i].font.bold=true
        rækker.items[i].font.name="Calibri" 
      }

      // Loop over celler
      var celler=rækker.items[i].cells
      celler.load('items')
      await context.sync()
      for (var k=0; k<celler.items.length; k++) {
        // Sætter padding
        celler.items[k].setCellPadding("Top",3)
        celler.items[k].setCellPadding("Bottom",3)
        // kolonnebredde
        if (k==0) {
          celler.items[k].columnWidth=248 
        }
        if (k>=1) {
          celler.items[k].columnWidth=61
        }
        
        // Højrestiller kolonneoverskrifter til højre hvis kun tal
        if (i==0&k>0) {
          const årstal=/^\d{4}$/;
          if (årstal.test(celler.items[k].value)==true) {
            celler.items[k].horizontalAlignment="Right"
          } else {
            celler.items[k].horizontalAlignment="Centered"
          }
        }
        // Højrestiller kolonne > 1 og række > 1
        if (i>0&k>0) {
          celler.items[k].horizontalAlignment="Right" 
        } 
        if (k==0&i>0&i<rækker.items.length-1) {
          celler.items[k].setCellPadding("Left",10)
        }
        // Styler i alt-rækker
        if (celler.items[k].value.slice(0,5)=="I alt") {
          console.log(celler.items[k].value.slice(0,5))
          rækker.items[i].shadingColor="#DDEBF7"
          rækker.items[i].font.bold=true
          celler.items[k].setCellPadding("Top",3)
          celler.items[k].setCellPadding("Bottom",3)
          var borderLocation = Word.BorderLocation.top;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({color:"#808080",width:1, type:'Single'})
          var borderLocation = Word.BorderLocation.bottom;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({color:"#808080",width:1, type:'Single'})
        }
        // Styler "servicerammen"
        if (celler.items[k].value.slice(0,13)=="Servicerammen"|celler.items[k].value.slice(0,22)=="Uden for servicerammen") {
          console.log(celler.items[k].value.slice(0,5))
          rækker.items[i].shadingColor="#BDD7EE"
          rækker.items[i].font.bold=true
          celler.items[k].setCellPadding("Top",3)
          celler.items[k].setCellPadding("Bottom",3)
          celler.items[k].setCellPadding("Left",5)
          var borderLocation = Word.BorderLocation.top;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({color:"#808080",width:1, type:'Single'})
          var borderLocation = Word.BorderLocation.bottom;
          var border = rækker.items[i].getBorder(borderLocation);
          border.set({color:"#808080",width:1, type:'Single'})
        }

      }
    }
    await context.sync();
  })
}

export async function replaceWordsWithLinks() {
  return Word.run(async (context) => { 
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const text = body.text;
      const regex = /(http(s)?:\/\/.)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/gi;
      const matches = text.match(regex);

      if (matches) {
          for (const match of matches) {
              const searchResults = body.search(match, { matchCase: false, matchWholeWord: true });
              context.load(searchResults, 'items');
              await context.sync();

              for (const result of searchResults.items) {
                  result.insertHyperlink(match, match, 'DisplayText');
              }
          }
      }
  }).catch(function (error) {
      console.log(error.message);
  });
}


// To run the function, call replaceWordsWithLinks()
replaceWordsWithLinks();