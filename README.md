# Liturgie → PowerPoint — webapp

Upload een liturgie (.docx), controleer een paar velden, download de kant-en-klare
dienst-PowerPoint. Dit is de webversie van alles wat in de chat is uitgewerkt.

## Wat is er nieuw in v6 (bugfixes na een doorlichting van v5)

- **Herkenning kerntekst / thema / schriftlezing** (`liturgie_parser.py`): niet meer alleen op het
  label, maar ook op de inhoud. Een bijbelverwijzing is nooit een thema; tekst zonder bijbelverwijzing
  is nooit een kerntekst. Herkent nu o.a. "Tekstkeuze: Johannes 3: 16", "Preek: Genade & waarheid",
  "Thema: Johannes 3: 16", "Bijbelgedeelte: Johannes 3: 1-21 (kerntekst: vers 16)",
  "Schriftlezing en tekst: ...", "Kerntekst: 'citaat' (Psalm 119: 105)" en "Thema (Johannes 3: 16)".
  Een kerntekst zonder boeknaam ("Tekst 2:13a 'Onderwerp u dan...'" na "Schriftlezing: 1 Petr. 2:11-25")
  krijgt het boek van de schriftlezing (1 Petrus 2: 13a) en wordt niet meer als tweede schriftlezing
  gezien. Een lezing als "1 Petr. 2:11-25 en 3:8-12" wordt twee lezingen (2:11-25 en 3:8-12), en die
  kunnen dan wel automatisch met tekst gevuld worden (samen was dat een lege tekst-dia).
  Staat het toch verkeerd om, dan wisselt de knop "Thema en kerntekst wisselen" op het controlescherm
  ze om. Alles wat de app zelf verplaatst staat in de bevindingen na het downloaden.
- **Voorzang uit het rooster**: "WPs", "WLd" en "Gz" (zoals in het psalmrooster) werden niet herkend, dus
  ongeveer 1 op 5 voorzangen verdween stilletjes. Nu wel; de voorzang staat al op het controlescherm
  (dus bij een Weerklank-voorzang kun je meteen het zipje uploaden) en wordt bij het downloaden opnieuw
  opgezocht als je kerk of tijd nog wijzigt.
- **`&` en `<` in thema/organist/voorganger/kerntekst** gaven een corrupte pptx. Opgelost (XML-escape).
- **Aanvangstijd**: leeg geeft geen "Aanvang 00:00 uur" meer en een onleesbare tijd ("half tien")
  laat de app niet meer crashen.
- **3+ schriftlezingen** staan nu allemaal op het overzicht en in de orde van dienst.
- **Collectedoelen** werden alleen gevonden als de app vanuit de eigen map gestart was.
- **Weerklank-zipjes** met submappen, spaties of hoofdletters worden nu verwerkt; lukt een zipje toch
  niet, dan staat dat in de bevindingen (vroeger: stil genegeerd).
- **Liturgie zonder kopjes**: te veel liederen (>6) of te weinig (<5) geeft een melding; een lied dat
  wegvalt omdat een plek al gevuld is ook.
- **Jaarroosters**: de app meldt wanneer het rooster op is (nu eind 2026) in plaats van "niet gevonden".
- **Veiligheid/onderhoud**: geen debug-modus meer standaard (`FLASK_DEBUG=1` om hem aan te zetten), eigen
  geheime sleutel (`.secret_key`), unieke bestandsnamen, alleen `.docx` als upload, oude downloads worden
  na 30 dagen opgeruimd (`OUTPUT_BEWAARDAGEN`).
- **`python selftest.py`**: controleert het rooster tegen de bibliotheek, de herkenning (32 gevallen) en
  of `template_map.json` nog klopt met het sjabloon. Draai dit na elke wijziging en na nieuwe data.

Nog niet op te lossen in code: **Gezang 4 vers 2** staat niet in het bronbestand (het bestand gaat van
vers 1 naar "3."). Het psalmrooster vraagt "Gz 4: 2" op 13 december 2026 (09:00 GK); die dia moet zelf
aangeleverd worden. Het jaarrooster voor 2027 moet je zelf inlezen met `psalmrooster_parser.py` en
`collecterooster_parser.py`.

## Snel starten (lokaal testen)

```bash
pip install -r requirements.txt
sudo apt install libreoffice tesseract-ocr poppler-utils   # alleen nodig als je zelf opnieuw wilt indexeren
python app.py
```

Ga naar `http://localhost:5000`.

## Wat moet er in `data/` staan

Dit staat NIET in deze zip (te groot / persoonlijk), dat moet je zelf toevoegen:

```
data/
  service_profiles.json          <- keuze voor standaarddienst en de meegeleverde avondmaalsvormen
  templates/                     <- ingevulde PowerPoints die als profielbasis dienen
    avondmaal_ochtend_voorbeeld.pptx
    avondmaal_ochtend_tweede_voorbeeld.pptx
    avondmaal_avond_voorbeeld.pptx
  psalm_index.json              <- output van psalm_indexer.py
  gezang_index.json             <- output van gezang_indexer.py (Gezang 1-12)
  gezang_correcties.json        <- handmatig gecontroleerde indelingen (optioneel)
  template_map.json             <- zit al in deze zip
  sjabloon_avonddienst.pptx     <- jullie kerksjabloon
  assets/
    orgel_grotekerk.png
    orgel_maranathakerk.png
  render_cache/
    Psalm_001_OB/
      slide-01.jpg
      slide-02.jpg
      ...
    Psalm_002_OB/
      ...
    (een map per psalmbestand, met de voorgerenderde dia's erin)
    Gezang_1__Tien_geboden/       <- idem voor de 12 gezangen (gezang_indexer.py)
    ...
  bijbeltekst/                   <- bijbelteksten voor de schriftlezingen
    books.json                   (uit de losse kerkdia-txt-app)
    verses.json
    bible-HSV/ bible-NBV/ bible-NBV21/
      GEN/1.json ...             (één JSON per hoofdstuk)
  fonts/
    GOTHICB.TTF                 <- Century Gothic Bold, voor het opmeten van
                                   de tekstregels (zelfde bestanden als in
                                   kerkdia-txt)
```

`psalm_index.json` en de `render_cache`-map heb ik in deze conversatie al één keer
volledig gegenereerd voor de 150 psalmen — vraag me erom als je die losse data
(circa 260MB) ook wilt downloaden, dan hoef je `psalm_indexer.py` niet zelf te
draaien. Wil je het wél zelf (opnieuw) genereren, bijvoorbeeld omdat de bibliotheek
gewijzigd is:

```bash
python psalm_indexer.py /pad/naar/psalmen_map data/
```
Dat zet dan meteen `psalm_index.json` én de `render_cache/`-map goed neer.

Voor de gezangen (map met "Gezang 1, Tien geboden.pptx" enz.):

```bash
python gezang_indexer.py /pad/naar/gezangen_map data/
```
Dat schrijft `data/gezang_index.json` en `data/render_cache/Gezang_*` (en een
`review_needed_gezang.json` met wat je even zelf moet nakijken; een handmatig
gecontroleerde indeling zet je in `data/gezang_correcties.json`).

## Deployen op je homeserver

De simpelste, degelijke manier voor een homeserver:

1. Zet deze map (+ de aangevulde `data/`) op je server.
2. Draai de app achter een echte WSGI-server in plaats van Flask's ingebouwde
   dev-server (die waarschuwing die je in de terminal ziet is serieus bedoeld):
   ```bash
   pip install gunicorn
   gunicorn -w 2 -b 0.0.0.0:5000 app:app
   ```
3. Zet er een reverse proxy voor (bv. Caddy of nginx) als je het ook van
   buiten je eigen netwerk bereikbaar wilt maken, met HTTPS.
4. Verander `app.secret_key` in `app.py` in iets unieks.

## Belangrijke beperkingen (bewust, geen bugs)

- **Dienstprofielen**: op het controlescherm kies je de standaarddienst,
  de Avondmaalsdienst of de Avondmaalsvoortzetting. Voor de Avondmaalsdienst
  kiest de app intern het passende ochtendvoorbeeld op basis van de aanvangstijd.
  De meegeleverde PowerPoints zijn de basis; liederen, dienstgegevens,
  lezingen, kerntekst, thema en collectedoelen. De zichtbaarheid van vaste dia's
  blijft zoals in het bijbehorende voorbeeld. Nieuwe dienstvormen vragen om een
  voorbeeldpresentatie en een bijpassende profielindeling.
- **Meer liedmomenten**: de avondmaalsprofielen verwerken de liederen in
  documentvolgorde, tot acht plekken per profiel. De eerste en tweede
  morgendienst en de voortzettingsdienst hebben elk hun eigen dia-indeling.

- **Gezangen 1 t/m 12** (Tien geboden, Lofzang van Maria/Zacharias/Simeon,
  Gebed des Heeren, Twaalf artikelen 1 en 2, Bedezang voor de predicatie,
  Morgenzang, Bedezang voor/Dankzang na het eten, Avondzang) staan net als de
  psalmen in de eigen bibliotheek (`gezang_index.json` + `render_cache/Gezang_*`).
  "Gezang 5: 1, 9" of "Gez. 5: 1, 9" wordt daaruit automatisch gevuld. Een
  gezang met "Wk"/"Weerklank" erbij, of een nummer hoger dan 12, blijft een
  Weerklank-lied (zipje-uploadvak). Opmerking: in de bronbestanden heeft
  Gezang 4 geen vers 2 (het tweede vers is daar als "3" genummerd) -- een
  liturgie met "Gezang 4: 2" geeft daarom een titel-placeholder en een melding
  in het dienstverslag.
- **Weerklank-psalmen en -liederen** (herkend aan "Wk", "Weerklank", of een
  Lied-regel / Gezang-regel met nummer boven 12) komen niet uit de eigen bibliotheek, maar uit
  weerklank.liedbundels.nu. Op het controlescherm verschijnt daarom per
  Weerklank-lied een uploadvak: zet daar het liedbundel-zipje neer (zoals
  de losse kerkdia-app die kreeg) en de app knipt de png's bij en plakt
  de dia's automatisch achter de titel-dia van dat lied (port van
  kerkdia, zie `weerklank.py`). Zonder zipje blijft er een titel-only
  placeholder-dia staan, klaar om de dia's zelf achter te plakken.
  Alleen echte psalmen (Psalm 1-150, Oude Berijming) en Gezang 1-12 worden
  uit de eigen bibliotheek gehaald; een psalmnummer dat niet in die bibliotheek zit
  geeft géén fout meer, maar ook een titel-only placeholder (genoteerd in
  het dienstverslag, zie onder).
- **Slim automatisch invullen uit het document** (`liturgie_parser.py`):
  - *Voorganger* alleen uit de titel (ds./prop./dr. + naam).
  - *Datum* ("zondag 27 sept", "27-09-2026"; zonder jaartal wordt het jaar
    gekozen waarin die dag op de juiste weekdag valt), *aanvangstijd*
    ("aanvang 10:00 uur", "9.30 uur"), *kerk* (Grote Kerk / Maranathakerk),
    *organist* en *type dienst* (aanvangstijd eerst, dan woorden als
    "zondagavond"; "Avondmaal"/"Morgenzang"/"Avondzang" tellen niet mee).
    Alles staat vooringevuld op het controlescherm.
  - *Thema* (ook "Onderwerp:") is nu een invulveld en dus aan te passen.
  - *Schriftlezing(en)*: labels als Schriftlezing, Lezing, Bijbelgedeelte,
    Bijbellezing, Schriftgedeelte; meerdere verwijzingen in 1 regel
    ("Genesis 1: 1-5 en Marcus 4: 35-41") worden aparte lezingen.
  - *Kerntekst*: Kerntekst, Preektekst, Preek:, of "Prediking/Verkondiging/
    Tekst" gevolgd door een bijbelverwijzing.
  - Bijbelverwijzingen: veel meer schrijfwijzen (Joh 3,16-21 · Johannes 3 vers
    16 t/m 21 · Lukas · Mar. · Heb. · Jac. · Tit. · "(HSV)" achteraan ·
    "(de verloren zoon)") en hele hoofdstukken ("Jesaja 53", "Psalm 23").
    "Johannes 3: 16 e.v." wordt niet meer stilletjes tot 1 vers ingekort.
  - Op het controlescherm laat een groen/oranje regel onder elke lezing en
    kerntekst direct zien of hij herkend is en of de tekst automatisch
    ingevuld wordt (`/parse_bijbel`).
  - Onder de datum staan de diensten uit het psalmrooster (kerk, tijd,
    voorzang) als knoppen, plus de collectedoelen (`/rooster_info`). Een klik
    vult kerk en aanvangstijd in.
- **Scherm na het downloaden**: de download start vanzelf en op dezelfde
  pagina staan de aandachtspunten (overgeslagen verzen, afgeleid type
  dienst, geen voorzang/collecte in het rooster, ...). Ze staan ook in het
  dienstverslag.
- **Versbereiken**: "Psalm 25: 1-3" of "1 t/m 3" = vers 1, 2 en 3. Ontbreekt
  een vers in de bibliotheek, dan worden de andere verzen wel geplakt.
- **Aanvangstijd** mag als 9.30, 9:30 uur, 10u of 10 worden getypt. Elke
  download gebruikt een eigen werkmap (twee mensen tegelijk kan nu wel).
- **Psalmen/liederen handmatig aanpassen**: op het controlescherm heeft elk
  onderdeel een knop "Wijzig" (of "Toevoegen" als er niets gevonden is). Typ
  de regel zelf, bv. `Psalm 19: 3, 4`, `Gezang 5: 1, 9` of `Psalm 43: 3 Wk`;
  de app controleert hem meteen (zelfde parser als bij het inlezen, via
  `/parse_regel`), meldt of hij uit de eigen bibliotheek komt of dat er
  verzen ontbreken, en laat bij een Weerklank-lied het zip-uploadvak zien.
  Een leeg veld laat het onderdeel weg; "Annuleer" zet de oorspronkelijke
  regel terug. Aanpassingen komen in het dienstverslag onder 'Bevindingen'.
- **Geen waarschuwingen in de webapp**: het controlescherm toont bewust géén
  waarschuwingen meer (die zorgden voor verwarring). Alles wordt in plaats
  daarvan bijgehouden in een DIENSTVERSLAG op schijf: na elke gegenereerde
  dienst maakt de app onder `output/verslagen/<datum>_<tijd>_<kerk>/` een
  `verslag.txt` aan (buiten de webapp om) met daarin alle dienstgegevens
  (kerk, type dienst, voorganger, organist, thema, vertaling, voorzang,
  collectedoelen), de liturgie zoals op het controlescherm, per
  schriftlezing/kerntekst of die automatisch gevuld is, en een lijst
  'Bevindingen': wat er niet automatisch kon (psalm niet in de bibliotheek,
  deels verzen, geen liedbundel-zipje bij een Weerklank-lied, enz.) -- of
  juist de melding dat alles goed ging. Ook een mislukte samenstelling
  krijgt zo'n verslag, met de reden erbij.
- **Bijbeltekst** wordt wél automatisch ingevuld voor schriftlezingen én
  kerntekst — met de bijbel-JSONs uit de losse kerkdia-txt-app
  (data/bijbeltekst/, vertaling kiezen op het controlescherm; standaard
  HSV). De dia's zien er precies zo uit als wat kerkdia-txt altijd
  maakte: eerst een titeldia met de verwijzing, dan tekstdia's met
  maximaal 8 regels en versnummers in superscript. Voor een SCHRIFTLEZING
  moeten het hele verzen zijn (een verwijzing met deels verzen zoals
  "Ezechiël 47: 9b+12" krijgt daar nog steeds een lege placeholder).
  Op de KERNTEKST-dia's worden ook deels-verwijzingen ingevuld: de
  genoemde hele verzen plus de a/b-delen, die automatisch op de
  zinsgrens dichtst bij het midden van het vers gesplitst worden
  (best effort — de kerk kan zulke delen zelf ook anders kiezen).
  Een onbekend boek of een regel die geen bijbelverwijzing is, krijgt
  nog steeds een lege placeholder-dia om zelf in te plakken.
  NB: dit gebruikt de teksten die de kerk zelf in kerkdia-txt had staan;
  zorg dat je voor de vertaling die je kiest een recht hebt om die tekst
  te projecteren.
- **Documentvormen**: alleen de tabelvorm en de losse-tekst-vorm die in de
  conversatie zijn tegengekomen worden herkend. Een heel ander soort
  liturgie-document moet apart aan `liturgie_parser.py` toegevoegd worden.
- **Eén sjabloon**: `template_map.json` is uitgezocht voor dit ene
  avonddienst-sjabloon (dat ook voor ochtenddiensten wordt hergebruikt door
  Wet/Geloofsbelijdenis te wisselen). Een structureel ander sjabloon heeft
  een eigen `template_map.json` nodig.

## Bestanden in dit pakket

| Bestand | Wat het doet |
|---|---|
| `app.py` | Flask-webapp: upload, controlescherm, download |
| `liturgie_parser.py` | Zet een liturgie-document om naar gestructureerde data |
| `assemble.py` | Voegt liturgie + psalm-index + sjabloon samen tot de PowerPoint |
| `pptx_surgery.py` | Generieke gereedschapskist voor pptx-bewerking |
| `psalm_indexer.py` | Bouwt `psalm_index.json` uit de ruwe psalmbibliotheek (OCR) |
| `gezang_indexer.py` | Bouwt `gezang_index.json` + render_cache uit de 12 gezang-pptx'en (OCR, met start-dia-methode voor onregelmatige gezangen) |
| `bijbeltekst.py` | Herkent bijbelverwijzingen en maakt er schriftlezing-dia's van (port van kerkdia-txt) |
| `weerklank.py` | Zet liedbundel-zipjes (liedbundels online) om naar dia-png's (port van de losse kerkdia-app) |
| `dienstverslag.py` | Schrijft na elke gegenereerde dienst een monitorverslag per dienst (buiten de webapp om) |
| `data/template_map.json` | Legt vast waar alles in het sjabloon staat |
| `data/assets/orgel_*.png` | Titeldia-foto's per kerkgebouw |
| `data/bijbeltekst/` | Bijbel-JSONs (HSV/NBV/NBV21) voor automatisch gevulde schriftlezingen |
| `templates/*.html` | De webpagina's |
| `selftest.py` | Zelfcontrole van rooster, herkenning en sjabloon |

Getest tegen 2 echte liturgieën uit de eerdere conversatie (tabelvorm en losse
tekst) — inclusief een volledige simulatie van upload tot download via
Flask's test-client, niet alleen "leest logisch".
