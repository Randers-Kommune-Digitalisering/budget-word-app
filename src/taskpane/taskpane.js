/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("insertTable").onclick = insertTable;
  }
});

export async function insertTable() {
  return Word.run(async (context) => {
    // https://www.youtube.com/watch?v=9u6MGqf1J_I
    const currentYear = new Date(Date.now()).getFullYear();
    const budgetperiode=[currentYear+1,currentYear+2,currentYear+3,currentYear+4];
    const overskrift=[""].concat(budgetperiode);
  
    const data = [
      overskrift,
      ["Indtægt", "-3,2", "-", "-", "-"],
      ["Budget", "3,1","0,1","0,1","0,1"],
      ["Nettoresultat", "-0,1","0,1","0,1","0,1"],
    ];
    const table = context.document.body.insertTable(5, 5, "Start", data);

    // context.document.body.insertParagraph("test", Word.InsertLocation.end);

    await context.sync();
  });
 
}

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    const organisation =
    [
        {
            "udvalg":"Økonomiudvalget",
            "bevillingsomr":["Administration","Politisk organisation"]
        },
        {
            "udvalg":"Børne- og familieudvalget",
            "bevillingsomr":["Børn","Familie"]
        },
        {
            "udvalg":"Beskæftigelsesudvalget",
            "bevillingsomr":["Beskæftigelse, integration og ydelser"]
        },
        {
            "udvalg":"Socialudvalget",
            "bevillingsomr":["Tilbud til børn, unge og voksne med særlige behov"]
        },
        {
            "udvalg":"Miljø- og teknikudvalget",
            "bevillingsomr":["Miljø og teknik, skattefinansieret","Miljø og teknik, takstfinansieret"]
        },
        {
            "udvalg":"Sundheds-, idræts- og kulturudvalget",
            "bevillingsomr":["Kultur og fritid","Sundhed"]
        },
        {
            "udvalg":"Omsorgsudvalget",
            "bevillingsomr":["Tilbud til ældre"]
        },
        {
            "udvalg":"Landdistriktsudvalget",
            "bevillingsomr":["Landdistrikt"]
        },
        {
            "udvalg":"Skole- og uddannelsesudvalget",
            "bevillingsomr":["Skole","Pædagogisk Psykologisk Rådgivning"]
        },
        {
            "udvalg":"Erhvervs- og planudvalget",
            "bevillingsomr":["Erhverv og plan"]
        }
    ]

    fetch('organisation.json')
    .then((response) => response.json())
    .then((json) => console.log(json));

    // insert a paragraph at the end of the document.

    for (var key in organisation) {
      if (organisation.hasOwnProperty(key)) {
        for (var key2 in organisation[key].bevillingsomr) {
          if (organisation[key].bevillingsomr.hasOwnProperty(key2)) {
            const tekst=organisation[key].udvalg + " - " + organisation[key].bevillingsomr[key2]
            context.document.body.insertParagraph(tekst, Word.InsertLocation.end); 
          }
        }
      }
    }
    //await context.sync()
    //context.document.save();
    //const paragraph2 = context.document.body.insertParagraph(organisation[1].udvalg, Word.InsertLocation.end);

    // change the paragraph color to blue.
    // paragraph.font.color = "blue";

    // Ind


    await context.sync();
  });
}
