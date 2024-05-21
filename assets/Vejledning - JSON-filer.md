# JSON-filer
## Indhold
1. [Introduktion](#introduktion)
    1. [*dokumenttype.json*](#dokumenttype)
    2. [*organisation.json*](#organisation)
2. [Redigering af JSON-filerne](#redigering)
    1. [GitHub](#github)
    2. [Deployment på Kubernetes](#kubernetes)

## 1. Introduktion <a id="introduktion"></a>
Strukturen i skabelonerne genereres ud fra to JSON-filer: 
- *dokumenttype.json* 
- *organisation.json*

*dokumenttype.json* beskriver de forskelle dokumenter appen er i stand til at generere, mens *organisation.json* beskriver hvordan skabelonerne skal udformes for specifikke  organisatoriske enheder (forvaltninger). Begge filer er beskrevet yderligere nedenfor. 

### 1.1 *dokumenttype.json* <a id="dokumenttype"></a>
## Beskrivelse af JSON-fil

#### Topniveau (Liste over dokumenttyper)
JSON-filen starter med en liste, hvor hvert element repræsenterer en dokumenttype.

#### Dokumenttype: Budgetopfølgning
- **Type**: `"Budgetopfølgning"`
- **Langt Navn**: `"Budgetopfølgning og forventet regnskab"`
- **Notatdetaljer**: En liste over detaljer i notatet
  - `"Sagsnavn:"`
  - `"Sagsnummer:"`
  - `"Skrevet af:"`
  - `"E-mail:"`
  - `"Forvaltning:"`
  - `"Dato:"`
  - `"Sendes til:"`

- **Undertype**: En liste over typer af budgetopfølgninger
  - `"31. marts"`
  - `"31. maj"`
  - `"30. september"`

- **Sektioner**: En liste over sektioner i dokumentet
  - `"Resumé"`
  - `"Bevilling"`
  - `"Anlæg"`
  - `"Bevillingsansøgninger"`

- **Undersektioner**: En liste af objekter, hvor hver undersektion indeholder kategorier som:
  - **Bevilling**:
    - `"Servicerammen"`
    - `"Indkomstoverførsler"`
    - `"Ældreboliger"`
    - `"Brugerfinansieret område"`
    - `"Centrale refusionsordninger mv."`
    - `"Aktivitetsbestemt medfinansiering"`
  - **Anlæg**:
    - `"Skattefinansierede anlæg"`
    - `"Ældreboliger - anlæg"`
    - `"Brugerfinansierede udgifter - anlæg"`
    - `"Jordforsyning"`
  - **Bevillingsansøgninger**:
    - `"Servicerammen"`
    - `"Indkomstoverførsler"`
    - `"Ældreboliger"`
    - `"Brugerfinansieret område"`
    - `"Centrale refusionsordninger mv."`
    - `"Aktivitetsbestemt medfinansiering"`
    - `"Skattefinansierede anlæg"`
    - `"Ældreboliger - anlæg"`
    - `"Brugerfinansierede udgifter - anlæg"`
    - `"Jordforsyning"`

- **Tabelindhold**: En liste over tabeller, hvor hver tabel indeholder:
  - **Type**: En typeidentifikator (f.eks. 0)
  - **Overskrifter**: En liste over kolonneoverskrifter, der varierer afhængigt af typen:
    - Type 0:
      - `"Service (mio. kr.)"`
      - `"Oprindeligt budget"`
      - `"Korrigeret budget ekskl. overførsler"`
      - `"Bevillings-ansøgninger"`
      - `"Korrigeret budget ekskl. overførsler inkl. bevillingsans."`
      - `"Forventet forbrug"`
      - `"Årets forventede resultat ekskl. overførsler"`
      - `"Overførsler"`
      - `"Årets forventede resultat inkl. overførsler"`
    - (Lignende struktur for andre typer: 1, 2, 3, 4, 5, 6, 7)

- **Standardfodnoter**: En liste over fodnoter, der gælder for tabellerne:
  - `"Note: Minus angiver et mindreforbrug/overskud i Årets forventede resultat og overførsler. Plus angiver et merforbrug/underskud."`
  - `"Note: Minus angiver indtægter, plus angiver udgifter."`

#### Dokumenttype: Budgetbemærkninger del 1
- **Type**: `"Budgetbemærkninger del 1"`
- **Langt Navn**: `"Del 1"`
- **Sektioner**: En liste over sektioner i dokumentet
  - `"1. Beskrivelse af området"`
  - `"2. Hovedtal"`
  - `"3. Aktuel økonomisk situation på drift og anlæg"`

- **Undersektioner**: En liste af lister, hvor hver undersektion indeholder:
  - Sektion 1: Ingen undersektioner
  - Sektion 2:
    - `"2.1 Drift"`
    - `"2.2 Anlæg"`
  - Sektion 3:
    - `"3.1 Aktuel status"`
    - `"3.2 Budgetaftalen"`

### 1.2 *organisation.json* <a id="organisation"></a>
#### Topniveau (Liste over udvalg)
JSON-filen starter med en liste, hvor hvert element repræsenterer et udvalg.

#### Udvalg
Hvert udvalg er repræsenteret som et objekt med følgende nøgle-værdi-par:
- `"udvalg"`: Navnet på udvalget (f.eks. "Økonomiudvalget").
- `"forkortelse"`: En forkortelse for udvalgets navn (f.eks. "ØU").

#### Dokumenter
For hvert udvalg kan der være en liste over dokumenter under nøglen `"dokumenter"`.
- Hvert dokument er et objekt med nøgler som:
  - `"navn"`: Navnet på dokumentet (f.eks. "Budgetopfølgning").
  - `"sektioner"`: En liste over sektioner identificeret ved numre (f.eks. `[0,1,2,3]`).

#### Undersektioner
Under hvert dokument kan der være undersektioner specificeret som en liste af objekter under nøglen `"undersektioner"`.
- Hver undersektion indeholder forskellige kategorier, som f.eks. `"bevilling"`, `"anlæg"`, og `"bevillingsansøgninger"`, hver med lister af nummererede elementer (f.eks. `[[0,3], [0]]`).

#### Bevillingsområde
Under dokumenterne kan der også være bevillingsområder specificeret som en liste af objekter under nøglen `"bevillingsområde"`.
- Hvert bevillingsområde indeholder:
  - `"navn"`: Navnet på bevillingsområdet (f.eks. "Administration").
  - `"tabeller"`: En liste over tabeller, hvor hver tabel er et objekt med nøgler som:
    - `"nr"`: Nummeret på tabellen.
    - `"navn"`: Navnet på tabellen.
    - `"beskrivelse"`: En beskrivelse af tabellen.
    - `"note"`: Noter vedrørende tabellen.
    - `"projekter"`: Hvorvidt en linje med projekter indgår (0/1)
    - `"typeKolonner"`: En typeidentifikator for kolonner (f.eks. 0).
    - `"rækker"`: En liste af rækkenavne.

#### Anlæg og Bevillingsansøgninger
Der er separate kategorier for anlæg og bevillingsansøgninger, hver indeholdende tabeller med lignende struktur som beskrevet ovenfor.

#### CustomTabeller
Endelig er der en sektion `"customTabeller"`, som indeholder brugerdefinerede tabeller med detaljeret information om deres placering, noter, indledende tekst, kolonner og rækker.


## 2. Redigering af JSON-filer <a id="redigering"></a>
### 2.1 GitHub <a id="github"></a>
#### Opret en branch på et GitHub repository

1. **Naviger til budgetappens repository**:
   - Gå til [https://github.com/Randers-Kommune-Digitalisering/budget-word-app](https://github.com/Randers-Kommune-Digitalisering/budget-word-app) og log ind på din konto.

2. **Opret en ny branch**:
   - Klik på dropdown-menuen, hvor der står `main` øverst til venstre i filvisningen.
   - Skriv navnet på din nye branch i tekstfeltet. Fx. `redigering - uke`.
   - Klik på `Create branch: [branch-navn] from main` for at oprette branchen.

#### Rediger en fil med den indbyggede editor

1. **Naviger til filen**:
   - Find den fil, du vil redigere i repository-strukturen, og klik på filens navn for at åbne den. *dokumenttype.json* og *organiastion.json* ligger her:
   [github.com/Randers-Kommune-Digitalisering/budget-word-app/assets](https://github.com/Randers-Kommune-Digitalisering/budget-word-app/assets)

2. **Rediger filen**:
   - Klik på blyant-ikonet (`Edit this file`) øverst til højre for at åbne filen i den indbyggede editor.
   - Foretag de ønskede ændringer i filen ved at redigere tekstindholdet.

#### Commit Ændringer

1. **Commit Ændringerne**:
   - Øverst til højre er en knap, der hedder `Commit changes`. Klik på den når du vil "gemme". 
   - Skriv en beskrivende commit-besked i feltet `Commit message`.
   - Vælg `Commit directly to the [branch-navn] branch` for at committe ændringerne direkte til din nye branch.
   - Klik på `Commit changes` for at gemme dine ændringer.

#### Lav et Pull Request

1. **Åbn Pull Request i fanen øvers**:
   - Klik på fanen `Pull requests` øverst i repository-visningen.
   - Klik på den grønne knap `New pull request`.

2. **Vælg Branches**:
   - Sørg for, at `base` branch (den du ønsker at merge dine ændringer ind i) er korrekt valgt (f.eks. `main`).
   - Sørg for, at `compare` branch er sat til den branch, du har lavet dine ændringer på.

3. **Opret Pull Request**:
   - Klik på den grønne knap `Create pull request`.
   - Skriv en titel og en beskrivelse for din pull request, der forklarer, hvad ændringerne indebærer.
   - Klik på `Create pull request` for at oprette pull requesten.


### 2.2 Deployment på Kubernetes <a id="kubernetes"></a>
... sker af sig selv med CI via Github Actions. Efter lidt tid kan du tjekke filerne er blevet opdateret på [budget-word-app.prototypes.randers.dk/assets/](https://budget-word-app.prototypes.randers.dk/assets/)