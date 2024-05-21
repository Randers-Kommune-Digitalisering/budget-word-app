# JSON-filer
## Indhold
1. [Introduktion](#introduktion)
    1. [*dokumenttype.json*](#dokumenttype)
    2. [*organisation.json*](#organisation)
2. [Redigering af JSON-filerne](#redigering)
    1. [GitHub](#github)
    2. [Deployment på Kubernetes](#kubernetes)

## 1. Introduktion <a id="introduktion"></a>
Strukturen 

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

### 2.2 Deployment på Kubernetes <a id="kubernetes"></a>