import { formatSelectedTable, rydAlt, rydSidehoved } from "./utils";
import { styleTable } from "./styleTable.js";

export async function fetchAssets(adr) {
  return Word.run(async (context) => {
    var response = await fetch(adr, { cache: "reload" });
    return response.json();
  });
}

export async function generateDocumentFromJSON(jsonData) {
    return Word.run(async (context) => { 
        let data;

        var tableStyles = await fetchAssets("https://localhost:3000/assets/"+"tableStyles.json");

        try {
            data = typeof jsonData === "string" ? JSON.parse(jsonData) : jsonData;
        } catch (e) {
            console.error("Invalid JSON input", e);
            return;
        }
        if (!Array.isArray(data)) {
            console.error("Input JSON must be an array of elements");
            return;
        }

        const body = context.document.body;
        for (let i = 0; i < data.length; i++) {
            /* Dokumentformattering */
            if (data[i].type === "opvarming") {
                if (data[i].hasOwnProperty("rydSidehoved") && data[i].rydSidehoved === true) {
                    await rydSidehoved();
                }
                if (data[i].hasOwnProperty("rydAlt") && data[i].rydAlt === true) {
                    await rydAlt();
                }
            }

            /* Sidehoved */
            if (data[i].type === "sidehoved") {
                const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary).insertParagraph(data[i].indhold, data[i].placering);
                
                if (data[i].hasOwnProperty("skrifttype")){
                    if (data[i].skrifttype.hasOwnProperty("størrelse")) {
                        header.font.size = data[i].skrifttype.størrelse;
                    } 
                    if (data[i].skrifttype.hasOwnProperty("kursiv")) {
                        header.font.italic = data[i].skrifttype.kursiv;
                    } 
                    if (data[i].skrifttype.hasOwnProperty("justering")) {
                        header.alignment = data[i].skrifttype.justering;
                    }
                }
            }
            /* Afsnit */
            if (data[i].type === "afsnit") {
                var indhold 
                data[i].indhold === null ? indhold = "" : indhold = data[i].indhold;
                const afsnit = body.insertParagraph(indhold, Word.InsertLocation.end);
                if (data[i].hasOwnProperty("style")) {
                    afsnit.style = data[i].style;               
                }
                if (data[i].hasOwnProperty("styleBuiltIn")) {
                    afsnit.styleBuiltIn = data[i].styleBuiltIn;               
                }
            }
            /* Tekst */
            if (data[i].type === "tekst") {
                const tekst = body.insertText(data[i].indhold, Word.InsertLocation.end);
                if (data[i].hasOwnProperty("style")) {
                    tekst.style = data[i].style;              
                } 
                if (data[i].hasOwnProperty("styleBuiltIn")) {
                    tekst.styleBuiltIn = data[i].styleBuiltIn;               
                }  
                if (data[i].hasOwnProperty("skrifttype")){
                    if (data[i].skrifttype.hasOwnProperty("kursiv")) {
                        tekst.font.italic = data[i].skrifttype.kursiv;
                    } 
                    if (data[i].skrifttype.hasOwnProperty("størrelse")) {
                        tekst.font.size = data[i].skrifttype.størrelse;
                    }     
                }
            }
            /* Tabeller */
            if (data[i].type === "tabel") {
                const tabel = body.insertTable(data[i].indhold.length, data[i].indhold[0].length, Word.InsertLocation.end, data[i].indhold);
                if (data[i].hasOwnProperty("styleAsSelectedTable") && (data[i].styleAsSelectedTable === true)) {
                    tabel.select();
                    await context.sync();
                    formatSelectedTable();
                }
                if (data[i].hasOwnProperty("style")) {
                    console.log("Kommer her");
                    await styleTable(context, tabel, data[i].style);
                }
            }
        }
        await context.sync();
    }); 
}
