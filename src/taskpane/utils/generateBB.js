import { rydAlt, rydSidehoved, fetchAssets } from "./utils.js";
import { tabelgenerator } from "./generateTable.js";


function tomlinje() {
    return {type:"afsnit", indhold:null, styleBuiltIn:"Normal"};
}


export async function generateBB(configurl, dokument, udvalg, bevillingsomraade) {

    const currentYear = new Date(Date.now()).getFullYear();
    const lastYear = currentYear - 1;
    const lastYear2 = currentYear - 2;
    const budgetperiodeÅr1 = currentYear + 1;
    const budgetperiodeÅr2 = currentYear + 2;
    const budgetperiodeÅr3 = currentYear + 3;
    const budgetperiodeÅr4 = currentYear + 4;
    const budgetperiode = budgetperiodeÅr1 + "-" + budgetperiodeÅr4;

    let data=[];   

    const udvalgsdata = await fetchAssets(configurl+udvalg+".json");
    var dokumentdata = await fetchAssets("https://localhost:3000/assets/"+"dokumenttype.json"); 
    var customStyle = await fetchAssets("https://localhost:3000/assets/"+"tableStyles.json");

    if (dokument == "Budgetbemærkninger del 1 - ny") {
      dokumentdata = dokumentdata.filter((obj) => obj.type == dokument);
      dokumentdata = dokumentdata[0];

      /* Rydder standardskabelones sidehoved */
      data.push({type: "opvarming", rydSidehoved: true, rydAlt: true});

      /* Indsætter nyt sidehoved */
      for (let i = 0; i < dokumentdata.sidehoved.length; i++) {
        var parse = require("json-templates");
        var parseIndhold = parse(dokumentdata.sidehoved[i].indhold);
        data.push({
          type: "sidehoved", 
          indhold: parseIndhold({udvalg: udvalgsdata.udvalg, bevillingsomraade: bevillingsomraade}), 
          placering: dokumentdata.sidehoved[i].placering, 
          skrifttype:{
            størrelse: dokumentdata.sidehoved[i].skrifttype.størrelse, 
            justering: dokumentdata.sidehoved[i].skrifttype.justering
          }
        });
      }

      /* Indsætter sektioner og underliggende indhold */
      async function processSektioner(sektioner) {
        for (let i = 0; i < sektioner.length; i++) {
          data.push({type: "afsnit", indhold: sektioner[i].overskrift, style: sektioner[i].styling});
          if (sektioner[i].hasOwnProperty("tabel") && Array.isArray(sektioner[i].tabel)) {
            for (let j = 0; j < sektioner[i].tabel.length; j++) {
              var tabeldata = await tabelgenerator(sektioner[i].tabel[j].type, dokument, dokumentdata, udvalgsdata, bevillingsomraade)
              
              var tabel = dokumentdata.tabeller.filter((obj) => obj.type == sektioner[i].tabel[j].type)[0];

              var parse = require("json-templates")
              var parseBeskrivelse = parse(tabel.beskrivelse)

              data.push({
                type: "tabel", 
                titel: tabel.titel,
                beskrivelse: parseBeskrivelse({fra:budgetperiodeÅr1, til:budgetperiodeÅr4}),
                indhold: tabeldata, 
                styleAsSelectedTable: sektioner[i].tabel[j].styleAsSelectedTable, 
                style: customStyle.filter(obj => obj.name == sektioner[i].tabel[j].style)[0]
              })

              if (tabel.hasOwnProperty("note") && tabel.note !== "") {
                data.push({type: "tekst", indhold: tabel.note, skrifttype:{størrelse: tabel.noteStyle.størrelse, kursiv: tabel.noteStyle.kursiv}});
              } 
            }
          }
          data.push(tomlinje());
          if (sektioner[i].hasOwnProperty("sektioner") && Array.isArray(sektioner[i].sektioner)) {
            await processSektioner(sektioner[i].sektioner);
          }
        }
      }
      await processSektioner(dokumentdata.sektioner);

      /* JSON til at generere dokumentet */
      console.log(JSON.stringify(data));
      return data;
    }
} 
