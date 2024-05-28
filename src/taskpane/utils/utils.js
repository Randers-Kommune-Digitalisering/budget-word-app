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
              if (celler.items[k].value.slice(0,5)=="I alt"||celler.items[k].value.slice(-5)=="i alt") {
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