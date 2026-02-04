export async function tabelgenerator(type, dokument, dokumentdata, udvalg, bevillingsomraade, inputdata) {
    let tabeldata = [];

    const currentYear = new Date(Date.now()).getFullYear();
    const lastYear = currentYear - 1;
    const lastYear2 = currentYear - 2;
    const budgetperiodeÅr1 = currentYear + 1;
    const budgetperiodeÅr2 = currentYear + 2;
    const budgetperiodeÅr3 = currentYear + 3;
    const budgetperiodeÅr4 = currentYear + 4;
    const budgetperiode = budgetperiodeÅr1 + "-" + budgetperiodeÅr4;

    if (type === "bb01") {
        var tabel = dokumentdata.tabeller.filter((obj) => obj.type == type);
        tabel = tabel[0];

        var parse = require("json-templates");
        var parseKolonner = parse(tabel.kolonner);

        tabeldata.push(parseKolonner({
            bevillingsomraadeParse: bevillingsomraade, aar1: budgetperiodeÅr1, aar2: budgetperiodeÅr2, aar3: budgetperiodeÅr3, aar4: budgetperiodeÅr4
        }));

        /* Servicerammen */
        tabeldata.push(["Servicerammen", "", "", "", ""]);

        var udvalg = udvalg.bevillingsområde.filter((obj) => obj.navn == bevillingsomraade);
        udvalg = udvalg[0];
        var rowsFullArray = udvalg.drift[0].servicerammen;
        var rows = rowsFullArray.map(subArray => subArray[0])
        for (let i = 0; i < rows.length; i++) {
            tabeldata.push([rows[i], "", "", "", ""]);
        }
 
        /* Uden for servicerammen */
        const driftObj = udvalg.drift[0];
        const hasOtherKeys = Object.keys(driftObj).some(key => key !== "servicerammen");
        if (hasOtherKeys) {
            tabeldata.push(["Uden for servicerammen", "", "", "", ""]);
        }

        Object.keys(driftObj).forEach(key => {
            console.log("Nuværende nøgle:", key);
            if (key !== "servicerammen" & key !== "projekter") {
                console.log(driftObj[key])
                var rowsFullArray = driftObj[key];
                var rows = rowsFullArray.map(subArray => subArray[0])
                for (let i = 0; i < rows.length; i++) {
                    tabeldata.push([rows[i], "", "", "", ""]);
                }
            }
        });

        tabeldata.push(["I alt", "", "", "", ""]);

    }
    if (type === "testtabel") {
        var tabel = dokumentdata.tabeller.filter((obj) => obj.type == type);
        tabel = tabel[0];

        tabeldata.push(tabel.kolonner);
        tabeldata.push(["Data 1","Data 2","Data 3"]);
        tabeldata.push(["Data 4","Data 5","Data 6"]);
        tabeldata.push(["Subtotal","Subtotal 2","Subtotal 3"]);
        tabeldata.push(["Data 7","Data 8","Data 9"]);
        tabeldata.push(["I alt 7","I alt 8","I alt 9"]);
        

    }   

    console.log("Tabeldata genereret:", tabeldata);
    return tabeldata;
}