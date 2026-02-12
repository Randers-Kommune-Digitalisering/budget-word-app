import {fetchAssets} from "./utils.js";
import {tabelgenerator} from "./generateTable.js";
import {tomlinje} from "./jsonUtils.js";
import { configurl, tableStylesurl, dokumenttypeurl, budgetperiodeÅr1, budgetperiodeÅr4 } from "./constants.js";

function getAllKeys(obj) {
  let keys = [];
  if (Array.isArray(obj)) {
    for (const item of obj) {
      keys = keys.concat(getAllKeys(item));
    }
  } else if (typeof obj === "object" && obj !== null) {
    for (const key in obj) {
      keys.push(key);
      keys = keys.concat(getAllKeys(obj[key]));
    }
  }
  return keys;
}

export async function generateBO(dokument, udvalg, dato) {

    let data=[];   

    const udvalgsdata = await fetchAssets(configurl+udvalg+".json");
    var dokumentdata = await fetchAssets(dokumenttypeurl+"dokumenttype.json"); 
    var customStyle = await fetchAssets(tableStylesurl+"tableStyles.json");

    if (dokument == "Budgetopfølgning - ny") {
        dokumentdata = dokumentdata.filter((obj) => obj.type == dokument);
        dokumentdata = dokumentdata[0];

        /* Rydder standardskabelones sidehoved */
        data.push({type: "opvarming", rydSidehoved: true, rydAlt: true});

        /* Indsætter sektioner og underliggende indhold */
        async function processSektioner(sektioner, bevillingsomraadeNavn = null) {

        /* Udtrækker keys fra udvalgsdata for at kunne styre, hvilke sektioner der skal inkluderes i dokumentet */
        var test = udvalgsdata.bevillingsområde.filter((obj) => obj.navn == bevillingsomraadeNavn);
        var allKeys = getAllKeys(test)

        for (let i = 0; i < sektioner.length; i++) {
            if (sektioner[i].hasOwnProperty("identifier") && !allKeys.includes(sektioner[i].identifier)) {
                continue; 
            } else {
                var parse = require("json-templates") 
                var indholdParse = parse(sektioner[i].overskrift);
                data.push({
                    type: "afsnit", 
                    indhold: indholdParse({udvalg: udvalgsdata.udvalg, dato: dato, bevillingsomraade: bevillingsomraadeNavn}), 
                    style: sektioner[i].styling
                });
            
                if (sektioner[i].hasOwnProperty("standardtekst") && sektioner[i].standardtekst !== "") {
                    data.push({type: "tekst", indhold: sektioner[i].standardtekst, styleBuiltIn: "Normal"});
                } 
                if (sektioner[i].hasOwnProperty("tabel") && Array.isArray(sektioner[i].tabel)) {
                    for (let j = 0; j < sektioner[i].tabel.length; j++) {
                    var tabeldata = await tabelgenerator(sektioner[i].tabel[j].type, dokument, dokumentdata, udvalgsdata, bevillingsomraadeNavn)
                    
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
                        data.push(tomlinje());
                    } 
                    }
                }
                data.push(tomlinje());
                if (sektioner[i].hasOwnProperty("sektioner") && Array.isArray(sektioner[i].sektioner)) {
                    await processSektioner(sektioner[i].sektioner, bevillingsomraadeNavn);
                }
            }
        }
      }

    await processSektioner(dokumentdata.headersektioner);
    for (let i = 0; i < udvalgsdata.bevillingsområde.length; i++) {
        await processSektioner(dokumentdata.sektioner, udvalgsdata.bevillingsområde[i].navn );
    } 

    /* JSON til at generere dokumentet */
    console.log(JSON.stringify(data));
    return data;
    }
} 
