import { styleTable } from "./utils";

export async function fetchAssets(adr) {
  return Word.run(async (context) => {
    var response = await fetch(adr, { cache: "reload" });
    return response.json();
  });
}

function tomlinje() {
    return {type:"afsnit", indhold:null, styleBuiltIn:"Normal"};
}


export async function generateBB(configurl, del, udvalg, bevillingsomraade) {
    let data=[];   

    const udvalgsdata = await fetchAssets(configurl+udvalg+".json");
    var dokumentdata = await fetchAssets("https://localhost:3000/assets/"+"dokumenttype.json"); 

    if (del == "Budgetbemærkninger del 1 - ny") {
      console.log(dokumentdata);
      dokumentdata = dokumentdata.filter((obj) => obj.type == del);
      dokumentdata = dokumentdata[0];
      console.log(dokumentdata);

      /* Rekursiv funktion til at behandle sektioner og undersektioner */
      function processSektioner(sektioner) {
        for (let i = 0; i < sektioner.length; i++) {
          data.push({type: "afsnit", indhold: sektioner[i].overskrift, style: sektioner[i].styling});
          if (sektioner[i].hasOwnProperty("tabel") && Array.isArray(sektioner[i].tabel)) {
            for (let j = 0; j < sektioner[i].tabel.length; j++) {
              data.push({type: "tabel", indhold: sektioner[i].tabel[j].indhold, styleAsSelectedTable: sektioner[i].tabel[j].styleAsSelectedTable});
            }
          }
          data.push(tomlinje());
          if (sektioner[i].hasOwnProperty("sektioner") && Array.isArray(sektioner[i].sektioner)) {
            processSektioner(sektioner[i].sektioner);
          }
        }
      }
      processSektioner(dokumentdata.sektioner);

/*    data.push({type:"afsnit", indhold:"1. Beskrivelse af området", style:"Overskrift 2"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"2. Hovedtal", style:"Overskrift 2"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"2.1 Drift", style:"Overskrift 3"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"2.1.1 Servicerammen", style:"Overskrift 4"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"2.1.2 Uden for servicerammen", style:"Overskrift 4"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"2.2 Anlæg", style:"Overskrift 3"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"3 Aktuel økonomisk situation på drift og anlæg", style:"Overskrift 2"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"3.1 Aktuelt status", style:"Overskrift 3"});
      data.push(tomlinje());
      data.push({type:"afsnit", indhold:"3.2 Budgetaftalen", style:"Overskrift 3"});
      data.push(tomlinje()); */

        console.log(JSON.stringify(data));
        return data;
    }
} 
