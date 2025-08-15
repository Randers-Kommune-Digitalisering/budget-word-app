# Budgetværktøjer (budget-word-app)
## Før

```mermaid
%%{
  init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#3c3c3c',
      'primaryTextColor': '#fff',
      'primaryBorderColor': '#3c3c3c',
      'lineColor': '#F8B229',
      'secondaryColor': '#616161',
      'tertiaryColor': '#616161',
      'tertiaryTextColor': '#fff'
    }
  }
}%%

flowchart LR
    %% Inputs
    json1[📄 dokumenttype.json] 
    json2[📄 organisation.json] 
    user[👤 Brugerinput] 
    docm[📄 Startskabelon.docm]
    skabelon[📄 Skabelon]
    UI[<div style='text-align:left'>
    ⚙️ Budgetværktøjer - UI
          - taskpane.html
    </div>]
    motor[<div style='text-align:left'>
    ⚙️ Budgetværktøjer 
          - taskpane.js
          - taskpane.css
          - utils/data.js
          - utils/utils.js
    </div>]  

    json2 --> motor
    json2 --> UI
    json1 --> motor
    json1 --> UI

    user --> UI

    UI --> motor

    motor --> skabelon
    docm --> skabelon
``` 
## Nu
```mermaid

%%{
  init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#3c3c3c',
      'primaryTextColor': '#fff',
      'primaryBorderColor': '#3c3c3c',
      'lineColor': '#F8B229',
      'secondaryColor': '#616161',
      'tertiaryColor': '#616161',
      'tertiaryTextColor': '#fff'
    }
  }
}%%

flowchart LR
    %% Inputs
    json1[📄 dokumenttype.json] 
    json2[📄 organisation.json] 
    json3["📄 [udvalg].json"] 
    user[👤 Brugerinput] 
    docm[📄 Startskabelon.docm]
    skabelon[📄 Skabelon]
    UI[<div style='text-align:left'>
    ⚙️ Budgetværktøjer - UI
          - taskpane.html
    </div>]
    motor[<div style='text-align:left'>
    ⚙️ Budgetværktøjer 
          - taskpane.js
          - taskpane.css
          - utils/data.js
          - utils/utils.js
    </div>] 

    json3 --> motor
    json2 --> UI
    json1 --> motor
    json1 --> UI

    user --> UI

    UI --> motor

    motor --> skabelon
    docm --> skabelon
``` 
## To be
```mermaid
%%{
  init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#3c3c3c',
      'primaryTextColor': '#fff',
      'primaryBorderColor': '#3c3c3c',
      'lineColor': '#F8B229',
      'secondaryColor': '#616161',
      'tertiaryColor': '#616161',
      'tertiaryTextColor': '#fff'
    }
  }
}%%

flowchart LR
    %% Inputs
    json1[📄 dokumenttype.json] 
    json3["📄 [udvalg].json"] 
    json4[📄 Jobbeskrivelse.json] 
    user[👤 Brugerinput] 
    docm[📄 Startskabelon.docm]
    skabelon[📄 Skabelon]
    UI[<div style='text-align:left'>
    ⚙️ Budgetværktøjer - UI
          - taskpane.html
    </div>]
    fletter[<div style='text-align:left'>
    ⚙️ json-fletter
          - taskpane.js
    </div>]
    motor[<div style='text-align:left'>
    ⚙️ Budgetværktøjer 
          - taskpane.js
          - taskpane.css
          - utils/data.js
          - utils/utils.js
    </div>] 

    json3 --> fletter
    json3 --> UI
    json1 --> fletter
    json1 --> UI

    user --> UI

    UI --> fletter
    fletter --> json4
    json4 --> motor

    motor --> skabelon
    docm --> skabelon
``` 


https://learn.microsoft.com/en-us/javascript/api/word/word.documentproperties?view=word-js-preview#word-word-documentproperties-comments-member




```
[
    {
        "udvalg":"<Navn på udvalg>",
        "forkortelse":"<Forkortelse på udvalg>",
        "bevillingsområde":[
            {
                "navn"="<Navn på bevillingsområde #1>"
                "delomårde":[
                    "<span style="color:blue"><delområde #1></span>", 
                    "<delområde #2>",
                    "<delområde #3>",
                    ...
                ],                    
                "indkomstoverførsler":[],
                "ældreboliger":[],
                "brugerfinansieret":[],
                "centralerefusionsordninger":[]


            },
            {
                ...
            }
        ],
        "dokumenter":[
            {
                "navn":"Budgetopfølgning",
                "sektioner":[[]],
                "undersektioner":[
                    {
                        "bevilling":[[],[]],
                        "anlæg":[[]],
                        "bevillingsansøgninger":[[],[]]
                    }
                ],
                "customTabeller":[
                    {  
                        "navn":"ct1",
                        "placering":"Bevilling Administration Servicerammen Social og Arbejdsmarked",
                        "tabelnr":<nr>,
                        "kolonner":[
                            "<Kolonne #1>","<Kolonne #2>",...
                        ],
                        "rækker":[
                            "<Række #1>","<Række #2>",...
                        ]
                    }
                ]
            }
        ]

    }
]
```




## Office Javascript API (WordApi 1.5)


## VBA 



* [word.table.descr](https://learn.microsoft.com/en-us/office/vba/api/word.table.descr)
* [word.table.title](https://learn.microsoft.com/en-us/office/vba/api/word.table.title)