"""
dienstverslag.py
================
Schrijft na elke gegenereerde dienst een monitorverslag naar schijf: een
nieuwe map per dienst (met datum, tijd en kerk in de naam) met daarin
alle informatie van het controlescherm plus wat er goed ging of fout liep
tijdens het samenstellen. Dit gebeurt BEWUST buiten de webapp om: het
controlescherm blijft zo rustig en toont geen waarschuwingen meer (die
veroorzaakten verwarring), het verslag is voor wie het wil nalezen.

De map wordt aangemaakt onder de uitvoermap:
    <uitvoer>/verslagen/<datum>_<tijd>_<kerk>/verslag.txt
Naast de download in <uitvoer>/ blijft zo per dienst een dossier bewaard.

GEBRUIK (in app.py na assemble_dienst):

    verslag_pad = schrijf_dienstverslag(
        uitvoer_map=OUTPUT_DIR,
        bestandsnaam=bestandsnaam,
        liturgie=liturgie,
        kerk=kerk,
        datum=datum, aanvangstijd=aanvangstijd,
        voorganger=voorganger, organist=organist,
        bijbel_vertaling=bijbel_vertaling,
        voorzang_opgezocht=voorzang_opgezocht,
        collecte_opgezocht=collecte_opgezocht,
        collectedoelen=collectedoelen,
        weerklank_afbeeldingen=weerklank_afbeeldingen,
        bevindingen=bevindingen,
    )
"""
import os
import re
from datetime import datetime

from bijbeltekst import verwijzing_parsen, verwijzing_ophalen

# De vaste volgorde van de dienst-onderdelen (zelfde labels als de webpagina).
SECTIE_VOLGORDE = [
    ("voorzang", "Voorzang"),
    ("opening", "Opening"),
    ("na_wet", "Na wet/apostolicum"),
    ("voor_prediking", "Voor de prediking"),
    ("na_prediking", "Na de prediking"),
    ("slot", "Slot"),
]

_KERKNAAM = {"GK": "Grote Kerk", "MK": "Maranathakerk"}


def _lezing_status(tekst, vertaling):
    """Hoe een schriftlezing in het eindresultaat is terechtgekomen."""
    if re.search(r"(?:catechismus|\bcat\.?\s*zondag|heid\.?\s*cat\.?)", tekst, re.IGNORECASE):
        return "catechismusverwijzing (als referentie getoond; geen automatische Bijbeltekst)"
    v = verwijzing_parsen(tekst)
    if v is None:
        return "niet herkend als bijbelverwijzing (titel-dia + lege tekst-dia)"
    if not v["heel"]:
        return "deels verzen -- niet automatisch te vullen (titel-dia + lege tekst-dia)"
    if verwijzing_ophalen(vertaling, v):
        return "automatisch gevuld met bijbeltekst"
    return "tekst niet beschikbaar in deze vertaling (titel-dia + lege tekst-dia)"


def _kerntekst_status(tekst, vertaling):
    """Hoe de kerntekst in het eindresultaat is terechtgekomen."""
    if not tekst:
        return "geen kerntekst opgegeven"
    v = verwijzing_parsen(tekst)
    if v is None:
        return "niet herkend als bijbelverwijzing (placeholder blijft staan)"
    if verwijzing_ophalen(vertaling, v, met_delen=True):
        if v["heel"]:
            return "automatisch gevuld met bijbeltekst"
        return ("automatisch gevuld; deels verzen: de a/b-delen zijn op een "
                "zinsgrens gesplitst (best effort)")
    return "tekst niet beschikbaar in deze vertaling (placeholder blijft staan)"


def schrijf_dienstverslag(uitvoer_map, *, bestandsnaam, liturgie, kerk,
                          datum, aanvangstijd, voorganger, organist,
                          bijbel_vertaling, voorzang_opgezocht=False,
                          collecte_opgezocht=False, collectedoelen=None,
                          weerklank_afbeeldingen=None, weerklank_fouten=None,
                          bevindingen=None, mislukt_reden=None):
    """Schrijft het monitorverslag en geeft het pad terug."""
    secties = liturgie.get("secties") or {}
    weerklank_afbeeldingen = weerklank_afbeeldingen or {}
    weerklank_fouten = weerklank_fouten or {}
    bevindingen = list(bevindingen or [])

    dag = (datum or datetime.now().strftime("%Y-%m-%d")).replace("/", "-").strip()
    tijd = (aanvangstijd or "").strip().replace(":", "-") or "tijd-onbekend"
    map_pad = os.path.join(uitvoer_map, "verslagen", f"{dag}_{tijd}_{kerk}")
    os.makedirs(map_pad, exist_ok=True)
    verslag_pad = os.path.join(map_pad, "verslag.txt")

    lijnen = []
    l = lijnen.append

    l("=" * 60)
    l("DIENSTVERSLAG")
    l("=" * 60)
    if mislukt_reden:
        l("STATUS: MISLUKT -- de PowerPoint kon niet worden samengesteld.")
    l(f"Gegenereerd:   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    l(f"PowerPoint:    {bestandsnaam}")
    l("")

    l("DIENSTGEGEVENS")
    l("-" * 60)
    l(f"Datum dienst:  {dag}")
    l(f"Aanvangstijd:  {aanvangstijd or '(niet ingevuld)'}")
    l(f"Kerk:          {kerk} ({_KERKNAAM.get(kerk, '?')})")
    soorten = {"ochtend": "Ochtenddienst", "avond": "Avonddienst"}
    l(f"Type dienst:   {soorten.get(liturgie.get('type_dienst'), '(niet herkend)')}")
    l(f"Voorganger:    {voorganger or '(niet ingevuld)'}")
    l(f"Organist:      {organist or '(niet ingevuld)'}")
    l(f"Thema:         {liturgie.get('thema') or '(geen thema opgegeven)'}")
    l(f"Bijbelteksten: automatisch ingevuld uit de {bijbel_vertaling}")
    l("")

    l("LITURGIE (zoals op het controlescherm)")
    l("-" * 60)
    profiel_momenten = liturgie.get("dienstprofiel") != "standard"
    muziek = (liturgie.get("muziekmomenten") or []) if profiel_momenten else []
    if muziek:
        for i, regel in enumerate(muziek, 1):
            if not regel:
                l(f"Liedmoment {i}: (leeg; plek weggelaten)")
                continue
            nummer = str(regel.get("nummer", "0"))
            try:
                nummer = str(int(nummer))
            except ValueError:
                pass
            nummer += str(regel.get("suffix", ""))
            verzen = ", ".join(
                str(v) + ("her" if v in regel.get("herhaal", []) else "")
                for v in regel.get("verzen", []))
            regel_naam = f"{regel.get('soort', 'Lied')} {nummer}: {verzen}"
            label = regel.get("label")
            if label and label != "Gebed des Heeren":
                regel_naam = f"{label}: {regel_naam}"
            if regel.get("bundel") == "Wk":
                aantal = len(weerklank_afbeeldingen.get(str(i - 1)) or [])
                status = (f"{aantal} Weerklank-dia's toegevoegd" if aantal else
                          "alleen titel-dia; handmatig aanvullen als er geen zipje is")
                l(f"Liedmoment {i}: {regel_naam} (Weerklank) -- {status}")
            else:
                l(f"Liedmoment {i}: {regel_naam} -- automatisch gevuld uit de bibliotheek")
        l("")
    for naam, label in SECTIE_VOLGORDE:
        if profiel_momenten:
            break
        regel = secties.get(naam)
        if not regel:
            l(f"{label}: (staat niet in de liturgie -- weggelaten)")
            continue
        nummer = int(regel.get("nummer", 0))
        verzen = ", ".join(
            str(v) + ("her" if v in regel.get("herhaal", []) else "")
            for v in regel.get("verzen", []))
        regel_naam = f"{regel.get('soort', 'Lied')} {nummer}: {verzen}"
        if regel.get("bundel") == "Wk":
            aantal = len(weerklank_afbeeldingen.get(naam) or [])
            if aantal:
                l(f"{label}: {regel_naam} (Weerklank) -- titel-dia + "
                  f"{aantal} lieddia's uit het geuploade zipje")
            elif naam in weerklank_fouten:
                l(f"{label}: {regel_naam} (Weerklank) -- alleen een titel-dia "
                  f"(zipje kon niet verwerkt worden: {weerklank_fouten[naam]})")
            else:
                l(f"{label}: {regel_naam} (Weerklank) -- alleen een titel-dia "
                  "(geen zipje geupload; dia's zelf achterplakken)")
        else:
            bron = "gezangenbibliotheek" if regel.get("soort") == "Gezang" else "psalmbibliotheek"
            l(f"{label}: {regel_naam} -- automatisch gevuld uit de {bron}")
    l("")

    lezingen = liturgie.get("schriftlezingen") or []
    l("SCHRIFTLEZINGEN")
    l("-" * 60)
    if lezingen:
        for i, lezing in enumerate(lezingen, 1):
            l(f"{i}. {lezing}")
            l(f"   -> {_lezing_status(lezing, bijbel_vertaling)}")
    else:
        l("(geen schriftlezingen in de liturgie gevonden)")
    l("")
    l("KERNTEKST")
    l("-" * 60)
    l(liturgie.get("kerntekst") or "(geen kerntekst opgegeven)")
    l(f"-> {_kerntekst_status(liturgie.get('kerntekst'), bijbel_vertaling)}")
    l("")

    l("VOORZANG EN COLLECTE")
    l("-" * 60)
    if secties.get("voorzang"):
        bron = "uit het jaarrooster opgezocht" if voorzang_opgezocht else "uit de liturgie"
        l(f"Voorzang: {bron}")
    else:
        l("Voorzang: (niet ingevuld; alleen als de liturgie of het rooster "
          "die geeft, komt er een voorzang-dia)")
    if collecte_opgezocht:
        if collectedoelen:
            doelen = ", ".join(f"{i + 1}. {doel}" for i, doel in enumerate(collectedoelen))
            l(f"Collectedoelen ({datum}): {doelen}  (collecterooster, wijk 3-4)")
        else:
            l(f"Collectedoelen: geen doelen gevonden in het collecterooster "
              f"voor {datum}")
    else:
        l("Collectedoelen: niet opgezocht (geen datum ingevuld)")
    l("")

    l("BEVINDINGEN")
    l("-" * 60)
    if mislukt_reden:
        l(f"LET OP: {mislukt_reden}")
        l("Er is daarom geen dienst-PowerPoint gedownload.")
    elif bevindingen:
        for b in bevindingen:
            l(f"- {b}")
        l("")
        l("Let op: bovenstaande onderdelen zijn niet (volledig) automatisch "
          "gevuld. Kijk na of ze handmatig moeten worden aangevuld of "
          "geplakt.")
    else:
        l("Alles goed gegaan: de dienst-PowerPoint is volledig automatisch "
          "samengesteld, zonder ontbrekende onderdelen.")

    with open(verslag_pad, "w", encoding="utf-8") as f:
        f.write("\n".join(lijnen) + "\n")
    return verslag_pad
