"""
liturgie_parser.py
===================
Zet een geüpload liturgie-document om naar gestructureerde data die
assemble.py kan gebruiken. Herkent op dit moment 2 vormen (meer heeft
deze kerk tot nu toe niet aangeleverd -- een nieuwe vorm moet er apart
bij gebouwd worden, dit is GEEN generieke alles-kan-parser):

  1. TABELVORM  -- vaste labels als "Openingspsalm/-lied:", "Psalm/lied na
     lezing wet/apostolicum:". Dit is de meest betrouwbare vorm.
  2. LOSSE TEKST -- een simpele chronologische opsomming (zoals een
     doorgestuurd e-mailtje), zonder vaste labels. Minder betrouwbaar --
     leunt op woordherkenning ("Wetslezing", "Prediking", etc.) om te
     bepalen WAAR een psalmregel bij hoort.

BELANGRIJKE AANNAME: dit koppelt liturgie-onderdelen aan de secties uit
template_map.json via VASTE NAMEN (voorzang, opening, na_wet,
voor_prediking, na_prediking, slot). Als de kerk een liturgie met een
andere volgorde/indeling aanlevert, moet die mapping hier aangepast worden.

GEBRUIK:
    from liturgie_parser import parse_liturgie
    data = parse_liturgie(pad_naar_docx_of_tekst)
"""
import os
import re
import subprocess
from datetime import date

from bijbeltekst import begint_met_boeknaam, splits_verwijzingen, verwijzing_parsen


# ---------------------------------------------------------------------------
# Hulpfuncties voor het parsen van 1 psalmregel, zoals "Ps. 100: 1 en 4her"
# of "Psalm 68, 16 en 17" of "Wk 433: 1, 2 en 3"
#
# BELANGRIJK over de bundels:
#   OB  = de eigen psalmbibliotheek (Psalm 1-150, Oude Berijming 1773) --
#         daarvoor bestaan voorgerenderde dia's in psalm_index.json.
#   Wk  = de Weerklank-bundel (weerklank.liedbundels.nu) en álle andere
#         liederen/gezangen die NIET uit die psalmbibliotheek komen. Deze
#         worden nooit automatisch uit psalm_index.json gehaald: de app
#         zet er alleen een titel-placeholder-dia voor neer.
#
# In de liturgie komt Weerklank op veel manieren voor, allemaal herkend:
#   'Weerklank lied 192: 1 en 2'   'Weerklank Psalm 145: 4, 5'
#   'Weerklank Lied: 302: 8, 9'    'Wk 433: 1, 2 en 3'
#   'Psalm 43, 3, 4 en 5 Wk'       'Psalm 148: 1, 2 (Wk)'
# ---------------------------------------------------------------------------

# soort-woord waarmee een regel kan beginnen/bevatten, gevolgd door een nummer.
# 'ps'/'psalm' = psalm uit de eigen bibliotheek; alle andere woorden zijn
# liederen/gezangen uit de externe bundel.
# De eigen bibliotheek bevat Gezang 1 t/m 12 (Oude Berijming: Tien geboden,
# Lofzang van Maria, ... Avondzang). 'Gezang 5' zonder Wk-markering is dus dat
# gezang; 'Gezang 5 Wk' / 'Weerklank Gezang 5' blijft een Weerklank-lied, en
# een hoger nummer (Gezang 123) is niet uit de eigen bibliotheek.
AANTAL_EIGEN_GEZANGEN = 12

_PSALMREGEL_PATROON = re.compile(
    r"(?<![a-z])(w?ps\.?|psalm|lied|w?ld|gezang|w?gz\.?|w?gez\.?|wk)\s*:?\s*(\d{1,3})([a-z]?)\b",
    re.IGNORECASE)
# Regel zonder soortwoord: 'Weerklank 192: 1, 2' -> nummer = Liednummer
_WEERKLANK_ZONDER_SOORT = re.compile(
    r"^\s*(?:weerklank|wk)\s*:?\s*(?:lied\s*:?\s*)?(\d{1,3})\b(.*)$",
    re.IGNORECASE)


def parse_psalmregel(tekst):
    """Zet een tekstregel als 'Ps. 100: 1 en 4her' of 'Psalm 43, 3, 4 en 5 Wk'
    om naar {'bundel': 'OB'|'Wk', 'soort': 'Psalm'|'Lied'|'Gezang',
    'nummer': '100', 'verzen': [1,4], 'herhaal': [4]}. Geeft None als er geen
    psalmregel in herkend wordt.

    bundel is 'Wk' (=> titel-placeholder-dia, geen automatische psalmdia's)
    zodra de regel Weerklank is óf geen psalm/gezang (1-12) uit de eigen
    bibliotheek is:
    - expliciete markering: 'Wk', '(Wk)', het woord 'Weerklank', of een
      regel die met 'Ld/Lied'-soort begint;
    - óf het soortwoord is 'lied'/'ld'/'gezang'/'wk' (Liednummer, niet 1-150).
    Alleen een regel die écht een psalm (ps/psalm) of een eigen gezang
    (Gezang 1-12) is, zonder Weerklank-markering, wordt 'OB' en kan uit de
    eigen bibliotheek worden ingevoegd."""
    tekst = tekst.strip()
    if not tekst:
        return None
    is_wk = (bool(re.search(r'\bwk\b', tekst, re.IGNORECASE))
             or re.search(r'weerklank', tekst, re.IGNORECASE) is not None
             or tekst.lower().startswith('ld'))

    m = _PSALMREGEL_PATROON.search(tekst)
    if m:
        soort_woord = m.group(1).lower()
        if soort_woord[0] == 'w' and soort_woord != 'wk':
            # WPs / WLd / WGz (zoals in het psalmrooster) = Weerklank
            is_wk = True
            soort_woord = soort_woord[1:]
        nummer = m.group(2).zfill(3)
        suffix = m.group(3).upper()
        rest = tekst[m.end():]
    else:
        # geen psalm/lied-woord gevonden: alleen zinvol als er wél een
        # Weerklank-markering staat ('Weerklank 192: 1, 2')
        m2 = _WEERKLANK_ZONDER_SOORT.match(tekst)
        if not m2 or not is_wk:
            return None
        nummer = m2.group(1).zfill(3)
        suffix = ""
        rest = m2.group(2)
        soort_woord = "lied"

    # soort: alleen 'ps'/'psalm' is een psalm uit de eigen bibliotheek;
    # 'lied'/'ld'/'wk'/'weerklank-zonder-soort' zijn Lied-nummers (Weerklank-
    # bundel / Liedboek), 'gezang' een gezang. Zo wordt 'Wk 433' dus nooit
    # meer ten onrechte 'Psalm 433' genoemd.
    if soort_woord.startswith("ps"):
        soort = "Psalm"
    elif soort_woord.startswith("gez") or soort_woord.startswith("gz"):
        soort = "Gezang"
    else:
        soort = "Lied"

    # opmerkingen tussen haakjes met cijfers ('(zie ook Joh 3)', '(melodie 134)')
    # horen niet bij de verzen; '(Wk)' bevat geen cijfers en blijft ongemoeid
    rest = re.sub(r"\([^)]*\d[^)]*\)", " ", rest)

    # bereiken uitschrijven: 'Psalm 25: 1-3' / '1 t/m 3' = vers 1, 2 en 3
    # (vroeger werden alleen 1 en 3 herkend en ging vers 2 stilletjes verloren)
    def _bereik(m):
        a, b = int(m.group(1)), int(m.group(2))
        return ",".join(str(n) for n in range(a, b + 1)) if a < b <= a + 40 else m.group(0)
    rest = re.sub(r"(\d{1,3})\s*(?:-|\u2013|\u2014|t/?m\b|tot\s+en\s+met)\s*(\d{1,3})", _bereik, rest,
                  flags=re.IGNORECASE)

    # verzen: cijfers, gescheiden door komma's/'en', met optioneel 'her' erachter
    verzen = []
    herhaal = []
    for vm in re.finditer(r'(\d{1,3})(her)?', rest):
        v = int(vm.group(1))
        verzen.append(v)
        if vm.group(2):
            herhaal.append(v)

    if not verzen:
        return None

    # alleen échte psalmen (zonder Weerklank-markering) en de 12 gezangen uit
    # de eigen bibliotheek (Gezang 1-12, zonder Weerklank-markering) komen uit
    # de eigen dia-bibliotheek (psalm_index.json / gezang_index.json); al het
    # andere is Wk => titel-placeholder (of Weerklank-zipje)
    eigen_gezang = soort == "Gezang" and 1 <= int(nummer) <= AANTAL_EIGEN_GEZANGEN
    bundel = "OB" if ((soort == "Psalm" or eigen_gezang) and not is_wk and not suffix) else "Wk"
    return {
        'bundel': bundel,
        'soort': soort,
        'nummer': nummer,
        'suffix': suffix,
        'verzen': verzen,
        'herhaal': herhaal,
        'ruwe_tekst': tekst,
    }


def verzen_label(parsed):
    """Nette weergave voor een titeldia, bv. 'Psalm 100: 1, 4' of
    'Psalm 43: 3, 4, 5 (Weerklank)'."""
    verzen_str = ", ".join(str(v) for v in parsed['verzen'])
    soort = parsed['soort'] if parsed['bundel'] == 'OB' else "Lied"
    label = f"{soort} {int(parsed['nummer'])}{parsed.get('suffix', '')}: {verzen_str}"
    if parsed['bundel'] == 'Wk':
        label += " (Weerklank)"
    return label


def parse_muziekmomenten(tekst):
    """Leest liedverwijzingen in documentvolgorde, ook als een liturgieregel
    meerdere momenten bevat (zoals '1e tafel ... , laatste tafel ...')."""
    momenten = []
    for regel in (tekst or "").splitlines():
        stukken = re.split(r"(?=\blaatste\s+tafel\b)", regel, flags=re.IGNORECASE)
        for stuk in stukken:
            label = "Bij voorbereiding tafel" if re.search(
                r"voorbereid(?:en|ing)?\s+(?:van\s+)?(?:de\s+)?tafel|voorbereiding\s+tafel",
                stuk, re.IGNORECASE) else None
            gebed_des_heeren = re.search(r"gebed\s+des\s+heeren", stuk, re.IGNORECASE)
            if gebed_des_heeren:
                verzen = re.search(r"gebed\s+des\s+heeren\D*([0-9][0-9, en]*)", stuk, re.IGNORECASE)
                if verzen:
                    parsed = parse_psalmregel("Gezang 5: " + verzen.group(1))
                    if parsed:
                        parsed["label"] = "Gebed des Heeren"
                        parsed["ruwe_tekst"] = stuk.strip()
                        momenten.append(parsed)
                continue
            parsed = parse_psalmregel(stuk)
            if parsed:
                parsed["label"] = label
                parsed["ruwe_tekst"] = stuk.strip()
                momenten.append(parsed)
    return momenten



# ---------------------------------------------------------------------------
# Kerntekst / thema / schriftlezing uit elkaar houden
#
# Vroeger bepaalde alleen het LABEL waar een waarde terechtkwam. Dat ging mis
# zodra een liturgie een ander woord gebruikte ("Tekstkeuze: Johannes 3: 16"
# = kerntekst, niet thema; "Preek: Genade & waarheid" = thema, geen kerntekst;
# "Bijbelgedeelte: Johannes 3: 1-21, kerntekst vers 16" = allebei). Nu beslist
# ook de INHOUD: een bijbelverwijzing is nooit een thema, en tekst zonder
# bijbelverwijzing is nooit een kerntekst.
# ---------------------------------------------------------------------------

_VERW_ZOEK = re.compile(
    r"(?<![\w])((?:[1-3]\s*)?[A-Za-z\u00c0-\u017f]{2,}\.?\s+\d{1,3}"
    r"(?:\s*[:,.]\s*\d{1,3}[a-z]?(?:\s*(?:-|\u2013|\u2014|t/?m|[,+]|en)\s*\d{1,3}[a-z]?)*)?)")
_KERN_WOORD = r"kerntekst|preektekst|hoofdtekst|tekstwoord|tekst|preek"
_LEZING_LABEL = re.compile(
    r"(?<![a-z])(?:schriftlezing(?:en)?|bijbelgedeelte|schriftgedeelte|bijbellezing|tekstlezing|lezing)")
_KERN_LABEL_STERK = re.compile(r"kerntekst|preektekst|hoofdtekst|tekstwoord")
_KERN_LABEL_ZWAK = re.compile(r"(?<![a-z])(?:tekst|preek|prediking|verkondiging)\b")
_THEMA_LABEL = re.compile(r"thema|onderwerp|tekstkeuze|preektitel|titel\s+(?:van\s+de\s+)?(?:preek|prediking)")


# Een kerntekst staat in een liturgie vaak ZONDER boeknaam, want het boek staat al bij de
# schriftlezing: "Tekst 2:13a 'Onderwerp u dan ...'" na "Schriftlezing: 1 Petr. 2:11-25".
_VERZEN_LIJST = r"\d{1,3}[a-z]?(?:\s*(?:-|\u2013|t/?m|,|\+|en)\s*\d{1,3}[a-z]?)*"
_KORT_HFD_VERS = rf"\d{{1,3}}\s*[:.]\s*{_VERZEN_LIJST}"
_KORT_VERS = rf"(?:vers|vss?\.?|v\.)\s*{_VERZEN_LIJST}"


def korte_verwijzing_begin(tekst, ook_kaal=False):
    """'2:13a 'Onderwerp u dan...'' -> '2:13a'; 'vers 13' -> 'vers 13'; anders None.
    Alleen 'hoofdstuk:vers' of 'vers N' telt; kale getallen ('13') alleen met ook_kaal=True
    (bij een stellig label als 'Kerntekst:')."""
    t = (tekst or "").strip()
    m = re.match(rf"^({_KORT_VERS}|{_KORT_HFD_VERS})(?![\w])", t, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    if ook_kaal:
        m = re.match(rf"^({_VERZEN_LIJST})(?![\w])", t, re.IGNORECASE)
        if m:
            return "vers " + m.group(1).strip()
    return None


def _volledige_kerntekst(kort, lezingen):
    """'2:13a' / 'vers 13' + ['1 Petr. 2:11-25', '1 Petr. 3:8-12'] -> '1 Petrus 2: 13a'.
    Het boek (en bij 'vers 13' het hoofdstuk) komt uit de schriftlezing; staat er een
    hoofdstuk bij, dan wordt de lezing gekozen waar dat hoofdstuk in valt."""
    m = re.match(r"^(?:(?:vers|vss?\.?|v\.)\s*)?(?:(\d{1,3})\s*[:.]\s*)?(\S.*)$", (kort or "").strip(), re.IGNORECASE)
    if not m:
        return None
    hoofdstuk, verzen = m.group(1), m.group(2).strip()
    kandidaten = [v for v in (verwijzing_parsen(l) for l in lezingen) if v]
    if not kandidaten:
        return None
    gekozen = kandidaten[0]
    if hoofdstuk:
        for v in kandidaten:
            if v["hoofdstuk_van"] <= int(hoofdstuk) <= (v["hoofdstuk_tm"] or v["hoofdstuk_van"]):
                gekozen = v
                break
    else:
        hoofdstuk = gekozen["hoofdstuk_van"]
    return f"{gekozen['boek_naam']} {hoofdstuk}: {verzen}"


def is_bijbelverwijzing(tekst):
    """True als de tekst met een bekende boeknaam + nummer begint ('Marcus 4: 35')."""
    return bool(tekst and tekst.strip()) and begint_met_boeknaam(tekst.strip())


def zoek_verwijzing(tekst):
    """Eerste bijbelverwijzing ergens binnen een langere tekst, of None."""
    for m in _VERW_ZOEK.finditer(tekst or ""):
        kandidaat = m.group(1).strip()
        if begint_met_boeknaam(kandidaat):
            return kandidaat
    return None


def _splits_verwijzing_aan_eind(tekst):
    """'Liefde (Johannes 3: 16)' / 'Liefde - Johannes 3: 16' -> ('Liefde', 'Johannes 3: 16')."""
    m = re.match(r"^(.*?\S)\s*[\(\[\-\u2013\u2014]\s*(.+?)\s*[\)\]]?\s*$", tekst or "")
    if m and is_bijbelverwijzing(m.group(2)) and not is_bijbelverwijzing(m.group(1)):
        return m.group(1).strip(), m.group(2).strip()
    return None


def label_soort(label, waarde=""):
    """Wat voor tekstveld is dit label? 'thema' | 'kerntekst' | 'lezing' | 'beide' | None.
    Zwakke labels (tekst, preek, prediking, verkondiging) tellen alleen mee als
    er echt een bijbelverwijzing achter staat."""
    l = (label or "").lower().strip(" :")
    if _THEMA_LABEL.search(l):
        return "thema"
    heeft_lezing = bool(_LEZING_LABEL.search(l)) and not re.search(r"wet|geloofs|apostol", l)
    heeft_kern = bool(_KERN_LABEL_STERK.search(l)) or bool(re.search(r"(?<![a-z])(?:kern|preek)?tekst\b", l))
    if heeft_lezing and heeft_kern:
        return "beide"
    if heeft_lezing:
        return "lezing"
    if _KERN_LABEL_STERK.search(l):
        return "kerntekst"
    if l == "bijbeltekst":
        # ambigu: een enkel vers is meestal de kerntekst, een bereik de lezing
        v = verwijzing_parsen(waarde)
        return "kerntekst" if v and v["vers_van"] is not None and v["vers_van"] == v["vers_tm"] and v["hoofdstuk_van"] == v["hoofdstuk_tm"] else "lezing"
    if _KERN_LABEL_ZWAK.search(l) and is_bijbelverwijzing(waarde):
        return "kerntekst"
    return None


def _zet_veld(result, sleutel, tekst):
    tekst = (tekst or "").strip(" '\u2018\u2019\"")
    if tekst and not result.get(sleutel):
        result[sleutel] = tekst


def _voeg_lezingen_toe(result, waarde):
    for verw in splits_verwijzingen(waarde):
        if verw not in result["schriftlezingen"]:
            result["schriftlezingen"].append(verw)


def _kern_uit_lezingwaarde(waarde):
    """'Johannes 3: 1-21, kerntekst vers 16' / 'Johannes 3: 1-21 (tekst: 3: 16)' /
    '1 Petr. 2:11-25 en 3:8-12; Tekst 2:13a 'Onderwerp u dan...'' ->
    (lezing-deel, kerntekst). Geen kerntekst-deel -> (waarde, None)."""
    m = re.search(rf"(?:^|[\s,;.(\-\u2013\u2014])(?:{_KERN_WOORD})\b\s*[:\-\u2013]?\s*(.+?)\s*\)?\s*$",
                  waarde, re.IGNORECASE)
    if not m or m.start() == 0:
        return waarde, None
    lezing_deel = re.sub(r"\s+en$", "", waarde[:m.start()].strip(" ,;.(-\u2013\u2014"), flags=re.IGNORECASE)
    kern = m.group(1).strip()
    if is_bijbelverwijzing(kern):
        return lezing_deel, kern
    kort = korte_verwijzing_begin(kern, ook_kaal=True)
    if kort:
        volledig = _volledige_kerntekst(kort, splits_verwijzingen(lezing_deel))
        return lezing_deel, volledig or kort
    return lezing_deel, kern or None


def verwerk_tekstveld(result, soort, waarde):
    """Zet 1 gelabelde waarde op de juiste plek (lezing / kerntekst / thema),
    op basis van label ÉN inhoud."""
    waarde = (waarde or "").strip()
    if not waarde:
        return
    waarde = re.sub(r"^(?:kern)?tekst\s*:\s*", "", waarde, flags=re.IGNORECASE) if soort in ("kerntekst",) else waarde

    if soort == "thema":
        m = re.search(r"thema\s*:\s*['\u2018\"]?([^'\u2019\"]+)", waarde, re.IGNORECASE)
        tekst = (m.group(1) if m else waarde).strip(" '\u2018\u2019\"")
        # 'Tekstkeuze: Johannes 3: 16; Thema: Liefde' -> allebei
        if m:
            voor = waarde[:m.start()].strip(" ;,-\u2013\u2014")
            if voor and is_bijbelverwijzing(voor):
                _zet_veld(result, "kerntekst", voor)
        if is_bijbelverwijzing(tekst) and not result.get("kerntekst"):
            _zet_veld(result, "kerntekst", tekst)          # 'thema' was eigenlijk de kerntekst
            return
        gesplitst = None if is_bijbelverwijzing(tekst) else _splits_verwijzing_aan_eind(tekst)
        if gesplitst:
            _zet_veld(result, "thema", gesplitst[0])
            _zet_veld(result, "kerntekst", gesplitst[1])
        else:
            _zet_veld(result, "thema", tekst)
        return

    if soort in ("lezing", "beide"):
        lezing_deel, kern = _kern_uit_lezingwaarde(waarde)
        if re.match(r"^tekst\s*:", lezing_deel, re.IGNORECASE):   # 'Tekst: Johannes 3: 16'
            _zet_veld(result, "kerntekst", re.sub(r"^tekst\s*:\s*", "", lezing_deel, flags=re.IGNORECASE))
            return
        if kern:
            _zet_veld(result, "kerntekst", kern)
        if lezing_deel:
            _voeg_lezingen_toe(result, lezing_deel)
            if soort == "beide" and not kern and len(splits_verwijzingen(lezing_deel)) == 1 \
                    and is_bijbelverwijzing(lezing_deel):
                _zet_veld(result, "kerntekst", lezing_deel)    # 'Lezing en tekst: Johannes 3: 16-21'
        return

    # soort == 'kerntekst'
    if is_bijbelverwijzing(waarde):
        _zet_veld(result, "kerntekst", waarde)
        return
    kort = korte_verwijzing_begin(waarde, ook_kaal=True)
    if kort:
        # boekloze kerntekst ("2:13a 'Onderwerp u dan...'"): het boek komt straks uit de
        # schriftlezing (zie herstel_kern_thema); de aangehaalde tekst laten we weg
        _zet_veld(result, "kerntekst", kort)
        return
    gesplitst = _splits_verwijzing_aan_eind(waarde)
    embedded = gesplitst[1] if gesplitst else zoek_verwijzing(waarde)
    if embedded:
        _zet_veld(result, "kerntekst", embedded)
    elif not result.get("thema"):
        _zet_veld(result, "thema", waarde)                # geen bijbelverwijzing -> is het thema
    else:
        _zet_veld(result, "kerntekst", waarde)            # (blijft placeholder; assemble meldt het)


def herstel_kern_thema(result):
    """Laatste vangnet na het parsen (of na handmatige aanpassing): thema en
    kerntekst verwisseld? Verplaatst ze en geeft een lijst meldingen terug."""
    meldingen = []
    thema = (result.get("thema") or "").strip() or None
    kern = (result.get("kerntekst") or "").strip() or None
    # kerntekst zonder boeknaam ('2:13a', 'vers 13'): boek en hoofdstuk uit de schriftlezing halen
    if kern and not is_bijbelverwijzing(kern) and korte_verwijzing_begin(kern, ook_kaal=True) == kern.strip():
        volledig = _volledige_kerntekst(kern, result.get("schriftlezingen") or [])
        if volledig:
            meldingen.append(f"Kerntekst '{kern}' had geen boek: aangevuld met het boek van de "
                             f"schriftlezing ({volledig}).")
            result["kerntekst"], kern = volledig, volledig
        else:
            meldingen.append(f"Kerntekst '{kern}' heeft geen boeknaam en er staat geen schriftlezing om "
                             "die uit over te nemen: vul de volledige verwijzing in (bv. 1 Petrus 2: 13).")
            return meldingen
    if thema and is_bijbelverwijzing(thema) and not (kern and is_bijbelverwijzing(kern)):
        result["kerntekst"], result["thema"] = thema, kern
        meldingen.append(f"Thema '{thema}' is een bijbelverwijzing: als kerntekst gebruikt"
                         + (f"; '{kern}' is nu het thema." if kern else "."))
    elif kern and not is_bijbelverwijzing(kern):
        emb = zoek_verwijzing(kern)
        if emb:
            result["kerntekst"] = emb
            meldingen.append(f"Kerntekst '{kern}': alleen de bijbelverwijzing '{emb}' is gebruikt.")
        elif not thema:
            result["thema"], result["kerntekst"] = kern, None
            meldingen.append(f"'{kern}' is geen bijbelverwijzing: als thema gebruikt, niet als kerntekst.")
    return meldingen


# ---------------------------------------------------------------------------
# Vorm 1: tabelvorm document lezen (rechtstreeks uit de Word-tabelcellen,
# via python-docx -- GEEN pandoc/tekst-tussenlaag, want pandoc breekt lange
# celwaarden over meerdere regels af en dat is onbetrouwbaar om te parsen)
# ---------------------------------------------------------------------------

TABEL_LABELS = {
    'voorzang':       ['voorzang'],
    'opening':        ['openingspsalm', 'openingslied'],
    'na_wet':         ['na lezing wet', 'wet/apostolicum'],
    'voor_prediking': ['voor de prediking'],
    'na_prediking':   ['na de prediking'],
    'slot':           ['slotpsalm', 'slotlied'],
}


def _label_match(label, kandidaten):
    label = label.lower()
    return any(k in label for k in kandidaten)


def parse_tabelvorm_docx(pad):
    from docx import Document
    doc = Document(pad)

    result = {'secties': {}, 'schriftlezingen': [], 'thema': None,
              'type_dienst': None, 'kerk': None, 'voorganger': None,
              'kerntekst': None, 'parse_bevindingen': []}

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 2:
                continue
            label, waarde = cells[0], cells[1]
            low_label = label.lower()
            low_waarde = waarde.lower()

            if not label or not waarde:
                continue
            if 'reeds ingevuld' in low_waarde:
                continue  # geen waarde beschikbaar, wordt later evt. via rooster ingevuld

            if 'type dienst' in low_label:
                if 'avond' in low_waarde:
                    result['type_dienst'] = 'avond'
                elif 'morgen' in low_waarde or 'ochtend' in low_waarde:
                    result['type_dienst'] = 'ochtend'
                continue

            # regels die bij een psalm-/liedplek horen nooit als lezing zien
            # ('Na lezing wet/apostolicum' bevat het woord 'lezing')
            hoort_bij_lied = any(_label_match(label, k) for k in TABEL_LABELS.values())

            if not hoort_bij_lied:
                soort = label_soort(label, waarde)
                if soort:
                    verwerk_tekstveld(result, soort, waarde)
                    continue

            for sectie, kandidaten in TABEL_LABELS.items():
                if _label_match(label, kandidaten):
                    parsed = parse_psalmregel(waarde)
                    if parsed:
                        result['secties'][sectie] = parsed
                    break

    return result


# ---------------------------------------------------------------------------
# Vorm 2: losse tekst (chronologische opsomming, zonder tabel)
# ---------------------------------------------------------------------------

# volgorde waarin we liturgische "ankers" verwachten, en welke sectie de
# psalmregel ERNA hoort te zijn
LOSSE_TEKST_ANKERS = [
    (re.compile(r'votum en groet', re.IGNORECASE), 'opening'),
    (re.compile(r'wetslezing|wet van god|lezing van de wet', re.IGNORECASE), 'na_wet'),
    (re.compile(r'collecte', re.IGNORECASE), 'voor_prediking'),
    (re.compile(r'prediking|verkondiging', re.IGNORECASE), 'na_prediking'),
    (re.compile(r'dankgebed|voorbede', re.IGNORECASE), 'slot'),
]

# Volgorde voor de "kale lijst"-vorm: een e-mail die gewoon de liederen
# achter elkaar opsomt, ZONDER liturgische ankerwoorden ("Wetslezing" etc)
# ertussen. Bij 6 liederen is de eerste de voorzang; bij 5 wordt de voorzang
# overgeslagen (die volgt dan uit het jaarrooster, zie psalmrooster_lookup.py).
KALE_LIJST_VOLGORDE_6 = ['voorzang', 'opening', 'na_wet', 'voor_prediking', 'na_prediking', 'slot']
KALE_LIJST_VOLGORDE_5 = ['opening', 'na_wet', 'voor_prediking', 'na_prediking', 'slot']

# Een regel die op een bijbeltekstverwijzing lijkt maar niet als psalmregel
# herkend is (dus geen "Schriftlezing:"-label ervoor, zoals bij een kale
# opsomming): "Filippenzen 1: 1 t/m 7", "Handelingen 16: 6-34",
# "1 Petr. 1: 3-9", etc. (het boeknummer van 1/2/3 Sam., Kor., Joh. ...
# mag voor de boeknaam staan)
BIJBELVERWIJZING_PATROON = re.compile(
    r'^(?:[1-3]\s+)?[A-Z][a-zA-Zëïü]+\.?\s+\d{1,3}\s*[:.]\s*\d{1,3}', re.UNICODE)


def parse_losse_tekst(tekst):
    result = {'secties': {}, 'schriftlezingen': [], 'thema': None,
              'type_dienst': None, 'kerk': None, 'voorganger': None,
              'kerntekst': None, 'parse_bevindingen': []}

    # (pandoc laat soms een '\\' aan het eind van een regel staan: weghalen)
    lines = [re.sub(r"\\([^\w\s])", r"\1", l).strip().rstrip('\\').strip()
             for l in tekst.splitlines() if l.strip()]

    # (type dienst wordt in parse_liturgie bepaald, via aanvangstijd en
    # titelwoorden -- niet meer op elk 'avond' in de tekst, want dat zag
    # Avondmaal/Avondzang ook voor een avonddienst aan)

    anker_gevonden = any(patroon.search(tekst) for patroon, _ in LOSSE_TEKST_ANKERS)

    # -- eerst: alle liederen op volgorde verzamelen, en alle overige
    #    bijzondere regels (schriftlezing/thema/preek) apart afvangen -----
    gevonden_liederen = []  # [(regelnummer, parsed), ...]
    for i, line in enumerate(lines):
        # 'Schriftlezing ...' met expliciet label. Er komen in de praktijk
        # meerdere schrijfwijzen voor: 'Schriftlezing: Genesis 1:1-10',
        # 'Schriftlezing 1: Genesis 1:1-10', 'Schriftlezing Genesis 1:1-10',
        # 'Lezing uit het Oude Testament: Jesaja 40: 1-11'. Pak de tekst ná
        # het label en haal daar de label-restjes af (nummertje + dubbele
        # punt, 'uit het Oude/Nieuwe Testament', enz.). De losse
        # 'lezing'-variant mag niet in 'Wetslezing'/'Geloofsbelijdenis'
        # matchen, vandaar de lookbehind -- én eist dat er echt een
        # bijbelverwijzing achter staat (zo wordt een kopregel als
        # 'Lezing van de Wet' niet voor een schriftlezing aangezien; die
        # hoort bij de wet, zie LOSSE_TEKST_ANKERS).
        m = re.search(r'(?:(schriftlezing(?:en)?|bijbelgedeelte|bijbellezing|schriftgedeelte|tekstlezing)'
                      r'|(?<![a-z])(?:lezing|lezen))\b(.*)', line, re.IGNORECASE)
        if m:
            expliciet_label = m.group(1) is not None
            rest = m.group(2).strip()
            afhaalpatronen = [
                r'^\d{1,2}\s*:\s*',                    # '1: Genesis ...'
                r'^:\s*',                                # ': Genesis ...'
                r'^(?:uit\s+het\s+)?(?:oude|nieuwe)\s+testament\s*:?\s*',
                r'^het\s+(?:oude|nieuwe)\s+testament\s*:?\s*',
            ]
            for _ in range(3):  # combinaties als '1: uit het O.T.: ...' aan
                for patroon in afhaalpatronen:
                    nieuw = re.sub(patroon, '', rest, count=1, flags=re.IGNORECASE)
                    if nieuw != rest:
                        rest = nieuw.strip()
                        break
                else:
                    break
            if rest:
                # 'Schriftlezing:' mag van alles bevatten (wordt later als
                # placeholder opgevangen); een kale 'Lezing ...' telt alleen
                # als er een echte verwijzing (Hoofdletter + nummer) achter staat
                ziet_eruit_als_bijbeltekst = bool(
                    re.match(r'^(?:[1-3]\s+)?[A-Z][A-Za-zëïü]*\.?\s+\d{1,3}', rest))
                if expliciet_label or ziet_eruit_als_bijbeltekst:
                    # 'Genesis 1: 1-5 en Marcus 4: 35-41' zijn twee lezingen;
                    # 'Schriftlezing en tekst: ...' / '..., kerntekst vers 16'
                    # levert ook de kerntekst op
                    soort = 'lezing'
                    mb = re.match(r'^(?:en|/|&)\s*(?:kern|preek)?tekst\s*:?\s*', rest, re.IGNORECASE)
                    if mb:
                        soort, rest = 'beide', rest[mb.end():]
                    verwerk_tekstveld(result, soort, rest)
            continue

        # kale bijbeltekstverwijzing zonder 'Schriftlezing:'-label ervoor
        # (alleen meetellen als er al minstens 1 lied gevonden is, om te
        # voorkomen dat een adresregel oid per ongeluk als lezing telt)
        if gevonden_liederen and not parse_psalmregel(line) and BIJBELVERWIJZING_PATROON.match(line) \
                and begint_met_boeknaam(line.strip()):
            verwerk_tekstveld(result, 'lezing', line.strip())
            continue

        # preek/kerntekst. Stellige labels (kerntekst, preektekst, ...)
        # nemen alles wat erachter staat; de bredere labels (tekst, prediking,
        # verkondiging) tellen alleen als er echt een bijbelverwijzing achter staat.
        # Wat erachter staat wordt op INHOUD beoordeeld (zie verwerk_tekstveld):
        # geen bijbelverwijzing -> dan is het het thema, geen kerntekst.
        # (een stellig label wordt altijd verwerkt, ook als er al een kerntekst is: staat
        # daar tekst zonder bijbelverwijzing, dan is dat het thema)
        # "Voor preek: Ps. 119:..." is een liedmoment, niet de preektekst.
        # De preektekst wordt expliciet aangegeven met "tekst vs ..." bij de lezing.
        m = re.search(r'(?<![a-z])(?:kerntekst|preektekst|hoofdtekst|tekstwoord)\b\s*(?:over\s+)?:\s*(.+)',
                      line, re.IGNORECASE)
        if m:
            verwerk_tekstveld(result, 'kerntekst', m.group(1).strip())
            continue
        if not result.get('kerntekst'):
            m = re.search(r'(?<![a-z])(?:tekst|prediking|verkondiging|kerntekst)\b\s*(?:over\s+)?[:\-\u2013]?\s*(\S.*)', line, re.IGNORECASE)
            if m and (begint_met_boeknaam(m.group(1)) or korte_verwijzing_begin(m.group(1))):
                verwerk_tekstveld(result, 'kerntekst', m.group(1).strip())
                continue

        # thema: alleen met een dubbele punt/streepje/aanhalingsteken erachter,
        # zodat 'thema' ergens middenin een psalmregel niet meer opgeslokt wordt
        if re.search(r"(?<![a-z])(?:thema|onderwerp)\b(?:[\]\)\*_]|\{\.[a-z\-]+\})*\s*(?:[:\-\u2013]|['\u2018\"])", line, re.IGNORECASE) and not result['thema']:
            # pak bij voorkeur de tekst tussen aanhalingstekens (rechte of
            # kromme), want ervoor kan nog markdown-opmaak-rommel staan
            # (pandoc zet *[Thema]{.underline}*: ervoor)
            qm = re.search(r"['\u2018\u2019]([^'\u2018\u2019]+)['\u2018\u2019]", line)
            if qm:
                thematekst = qm.group(1).strip()
                rest_buiten = (line[:qm.start()] + " " + line[qm.end():])
            else:
                # tekst NA het label (eerste dubbele punt/streepje), dus niet na de
                # laatste dubbele punt: 'Thema: Johannes 3: 16' bevat er twee
                lm = re.search(r"(?:thema|onderwerp)\b(?:[\]\)\*_]|\{\.[a-z\-]+\})*\s*[:\-\u2013]\s*(.+)$",
                               line, re.IGNORECASE)
                thematekst = lm.group(1).strip(" '\u2018\u2019\"") if lm else ""
                rest_buiten = ""
            verwerk_tekstveld(result, 'thema', thematekst)
            # 'Thema: 'Liefde' - Johannes 3: 16' -> de verwijzing is de kerntekst
            if not result.get('kerntekst') and rest_buiten:
                verw = zoek_verwijzing(rest_buiten)
                if verw:
                    result['kerntekst'] = verw
            continue

        parsed = parse_psalmregel(line)
        if parsed:
            gevonden_liederen.append((i, parsed))

    if anker_gevonden:
        # -- vorm A: ankerwoorden aanwezig (bv. "Wetslezing", "Collecte") --
        # de eerste psalmregel (vóór het eerste ankerwoord) is de voorzang,
        # daarna bepalen de ankers welke sectie een psalmregel toebehoort.
        eerste_anker_pos = min(
            (m.start() for patroon, _ in LOSSE_TEKST_ANKERS for m in [patroon.search(tekst)] if m),
            default=len(tekst))
        verwachte_sectie_index = 0
        for line in lines:
            for patroon, sectie in LOSSE_TEKST_ANKERS:
                if patroon.search(line):
                    verwachte_sectie_index = LOSSE_TEKST_ANKERS.index((patroon, sectie))
                    break
            parsed = parse_psalmregel(line)
            if parsed:
                if 'opening' not in result['secties'] and verwachte_sectie_index == 0 \
                   and tekst.find(line) < eerste_anker_pos:
                    sectie = 'voorzang' if 'voorzang' not in result['secties'] else LOSSE_TEKST_ANKERS[0][1]
                else:
                    sectie = LOSSE_TEKST_ANKERS[verwachte_sectie_index][1]
                if sectie not in result['secties']:
                    result['secties'][sectie] = parsed
                else:
                    result['parse_bevindingen'].append(
                        f"Lied '{parsed['ruwe_tekst']}' overgeslagen: het onderdeel '{sectie}' "
                        f"was al gevuld met '{result['secties'][sectie]['ruwe_tekst']}'. Controleer de liturgie.")
    else:
        # -- vorm B: kale lijst zonder ankerwoorden (bv. de-Kruijf-stijl) --
        # gewoon de gevonden liederen op volgorde aan de vaste plekken
        # hangen: 6 liederen = inclusief voorzang, 5 = voorzang overslaan.
        n = len(gevonden_liederen)
        if n == 6:
            volgorde = KALE_LIJST_VOLGORDE_6
        elif n == 5:
            volgorde = KALE_LIJST_VOLGORDE_5
        else:
            # onverwacht aantal (bv. maar 2 of 3 liederen) -- toch zo goed
            # mogelijk vullen door van achteren te vullen (slot het meest
            # betrouwbaar aanwezig), zodat er niet niks gebeurt.
            volgorde = KALE_LIJST_VOLGORDE_6[-n:] if n <= 6 else KALE_LIJST_VOLGORDE_6
        for (_, parsed), sectie in zip(gevonden_liederen, volgorde):
            result['secties'][sectie] = parsed
        if n > 6:
            over = ", ".join(p['ruwe_tekst'] for _, p in gevonden_liederen[6:])
            result['parse_bevindingen'].append(
                f"{n} liederen gevonden, maar er zijn maar 6 plekken: de laatste {n - 6} zijn niet gebruikt ({over}).")
        elif 0 < n < 5:
            result['parse_bevindingen'].append(
                f"Maar {n} {'lied' if n == 1 else 'liederen'} gevonden in een liturgie zonder vaste kopjes: "
                f"aan de laatste {n} {'plek' if n == 1 else 'plekken'} van de dienst gehangen. "
                "Controleer welk lied bij welk onderdeel hoort.")

    return result


# ---------------------------------------------------------------------------
# Vorm herkennen + hoofdfunctie
# ---------------------------------------------------------------------------

def docx_naar_tekst(pad):
    """Leest een .docx-bestand uit.

    Probeer eerst pandoc, maar fallback naar python-docx als pandoc niet
    beschikbaar is op dit systeem.
    """
    try:
        out = subprocess.run(
            ["pandoc", "-t", "markdown", "--wrap=none", pad],
            capture_output=True,
            text=True,
            check=True,
        )
        return out.stdout
    except (FileNotFoundError, subprocess.CalledProcessError):
        from docx import Document

        doc = Document(pad)
        tekstdelen = []

        for para in doc.paragraphs:
            if para.text.strip():
                tekstdelen.append(para.text.strip())

        for table in doc.tables:
            for row in table.rows:
                cellen = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if cellen:
                    tekstdelen.append(" | ".join(cellen))

        return "\n\n".join(tekstdelen)


def _docx_has_table(pad):
    try:
        from docx import Document
        doc = Document(pad)
        return len(doc.tables) > 0
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Voorganger uit de titel van het document
# ---------------------------------------------------------------------------

_TITEL_WOORD = re.compile(
    r"\b(ds|dr|drs|prof|prop|proponent|kand|kandidaat)\b\.?[ \t]+", re.IGNORECASE)
_TITEL_NORMAAL = {"ds": None,  # 'Ds.' zet de webapp zelf voor de naam
                  "dr": "Dr.", "drs": "Drs.", "prof": "Prof.",
                  "prop": "Prop.", "proponent": "Prop.",
                  "kand": "Kand.", "kandidaat": "Kand."}
_TUSSENVOEGSELS = {"van", "de", "der", "den", "ter", "ten", "te", "'t", "\u2019t",
                   "het", "la", "le", "du", "von", "op", "in", "aan"}
_INITIALEN = re.compile(r"^(?:[A-Z]\.?){1,5}$")


def _naam_na_titel(rest):
    """Pakt de naam die direct na 'ds.' staat: initialen, tussenvoegsels en
    hoofdletter-woorden, tot het eerste leesteken/cijfer/kleine woord."""
    tokens = []
    for woord in rest.split():
        eind_leesteken = woord[-1] in ",;:)-\u2013\u2014|"
        kaal = woord.strip(",;:()|-\u2013\u2014")
        if woord[0] in "(|[-\u2013\u2014" or not kaal or any(c.isdigit() for c in kaal):
            break
        if woord[-1] == ":":
            break  # 'Thema:' e.d. is een label, geen deel van de naam
        if kaal.lower() in _TUSSENVOEGSELS and tokens_heeft_geen_achternaam(tokens):
            tokens.append(kaal)
        elif _INITIALEN.match(kaal) or (kaal[0].isupper() and re.match(r"^[\w\u2019'.\-]+$", kaal)):
            tokens.append(kaal)
        else:
            break
        if eind_leesteken or len(tokens) >= 6:
            break
    # minstens één echte achternaam (dus niet alleen initialen/tussenvoegsels)
    if not any(t[0].isupper() and not _INITIALEN.match(t) for t in tokens):
        return None
    while tokens and tokens[-1].lower() in _TUSSENVOEGSELS:
        tokens.pop()
    return " ".join(tokens) or None


def tokens_heeft_geen_achternaam(tokens):
    """Tussenvoegsels (van, de, der...) horen vóór de achternaam; na een
    achternaam-woord duidt een klein woord op tekst die niet meer bij de naam hoort."""
    return not any(t[0].isupper() and not _INITIALEN.match(t) for t in tokens)


def voorganger_uit_regel(regel):
    """'Liturgie zondag 20 sept - ds. J. van der Linden' -> 'J. van der Linden'.
    Bij een proponent/kandidaat/dr. blijft de titel erbij ('Prop. Jansen'),
    zodat de webapp er dan geen 'Ds.' voor zet. None als er geen naam staat."""
    for m in _TITEL_WOORD.finditer(regel):
        naam = _naam_na_titel(regel[m.end():])
        if not naam:
            continue
        titel = _TITEL_NORMAAL[m.group(1).lower()]
        return f"{titel} {naam}" if titel else naam
    return None


def _titel_kandidaten(pad_of_tekst, is_bestandspad):
    """Regels die als 'titel van het document' tellen, in volgorde van belang:
    titel-/kopstijlen en de documenteigenschap 'titel', de eerste paar
    niet-lege regels (ook in de koptekst), en als laatste de bestandsnaam."""
    kandidaten = []
    if is_bestandspad and pad_of_tekst.lower().endswith('.docx'):
        try:
            from docx import Document
            doc = Document(pad_of_tekst)
            titel_prop = (doc.core_properties.title or "").strip()
            if titel_prop:
                kandidaten.append(titel_prop)
            for para in doc.paragraphs[:15]:
                stijl = (para.style.name if para.style is not None else "").lower()
                if para.text.strip() and re.match(r"(title|titel|heading|kop|subtitle|ondertitel)", stijl):
                    kandidaten.append(para.text.strip())
            for sec in doc.sections[:1]:
                for para in sec.header.paragraphs:
                    if para.text.strip():
                        kandidaten.append(para.text.strip())
            eerste = [p.text.strip() for p in doc.paragraphs if p.text.strip()][:3]
            kandidaten += eerste
        except Exception:
            pass
        kandidaten.append(os.path.splitext(os.path.basename(pad_of_tekst))[0].replace("_", " "))
    else:
        regels = [l.strip() for l in pad_of_tekst.splitlines() if l.strip()]
        kandidaten += regels[:3]
    return kandidaten


def vind_voorganger_in_titel(pad_of_tekst, is_bestandspad=True):
    for regel in _titel_kandidaten(pad_of_tekst, is_bestandspad):
        naam = voorganger_uit_regel(regel)
        if naam:
            return naam
    return None


# ---------------------------------------------------------------------------
# Dienstgegevens uit het document halen: datum, aanvangstijd, kerk, organist,
# type dienst. Alles wat hier gevonden wordt vult het controlescherm alvast
# in (en de gebruiker kan het daar nog aanpassen).
# ---------------------------------------------------------------------------

_MAANDEN = {"januari": 1, "jan": 1, "februari": 2, "feb": 2, "maart": 3, "mrt": 3,
            "april": 4, "apr": 4, "mei": 5, "juni": 6, "jun": 6, "juli": 7, "jul": 7,
            "augustus": 8, "aug": 8, "september": 9, "sept": 9, "sep": 9,
            "oktober": 10, "okt": 10, "november": 11, "nov": 11, "december": 12, "dec": 12}
_MAAND_ALT = "|".join(sorted(_MAANDEN, key=len, reverse=True))
_WEEKDAGEN = {"maandag": 0, "dinsdag": 1, "woensdag": 2, "donderdag": 3,
              "vrijdag": 4, "zaterdag": 5, "zondag": 6}
_DATUM_MET_MAAND = re.compile(
    rf"(?<![\d:.])(\d{{1,2}})(?:e|ste|de)?[\s\-]+({_MAAND_ALT})\b\.?(?:[\s,\-]+(\d{{4}}))?", re.IGNORECASE)
_DATUM_NUMERIEK = re.compile(r"(?<![\d:.])(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})(?![\d:])")
_WEEKDAG_ZOEK = re.compile(r"(?<![a-z])(" + "|".join(_WEEKDAGEN) + ")", re.IGNORECASE)


def _kies_jaar(dag, maand, weekdag, vandaag):
    """Zonder jaartal in het document: het jaar waarin de datum op de juiste
    weekdag valt (zondag!) en het dichtst bij vandaag ligt."""
    for filter_weekdag in ([weekdag, None] if weekdag is not None else [None]):
        kandidaten = []
        for jaar in (vandaag.year - 1, vandaag.year, vandaag.year + 1):
            try:
                d = date(jaar, maand, dag)
            except ValueError:
                continue
            if filter_weekdag is not None and d.weekday() != filter_weekdag:
                continue
            kandidaten.append(d)
        if kandidaten:
            return min(kandidaten, key=lambda d: abs((d - vandaag).days))
    return None


def vind_datum(regels, vandaag=None):
    """Eerste datum in de gegeven regels als 'YYYY-MM-DD' (of None)."""
    vandaag = vandaag or date.today()
    for regel in regels:
        weekdag_m = _WEEKDAG_ZOEK.search(regel)
        weekdag = _WEEKDAGEN[weekdag_m.group(1).lower()] if weekdag_m else None
        m = _DATUM_MET_MAAND.search(regel)
        if m:
            dag, maand = int(m.group(1)), _MAANDEN[m.group(2).lower()]
            if m.group(3):
                try:
                    return date(int(m.group(3)), maand, dag).isoformat()
                except ValueError:
                    continue
            d = _kies_jaar(dag, maand, weekdag, vandaag)
            if d:
                return d.isoformat()
        m = _DATUM_NUMERIEK.search(regel)
        if m:
            dag, maand, jaar = int(m.group(1)), int(m.group(2)), int(m.group(3))
            jaar = jaar + 2000 if jaar < 100 else jaar
            try:
                return date(jaar, maand, dag).isoformat()
            except ValueError:
                continue
    return None


_TIJD_MET_CONTEXT = re.compile(
    r"\b(?:aanvang(?:stijd)?|begin(?:tijd)?|tijdstip|tijd|om|vanaf|start)\b\s*:?\s*(?:ca\.?\s*)?"
    r"(\d{1,2})(?:\s*[:.]\s*([0-5]\d))?\s*(?:uur|u|h)?(?![\d:.,])", re.IGNORECASE)
_TIJD_MET_UUR = re.compile(r"(?<![\d:.,])(\d{1,2})(?:\s*[:.]\s*([0-5]\d))?\s*(?:uur|u)\b", re.IGNORECASE)
_TIJD_LOS = re.compile(r"(?<![\d:.,])(\d{1,2})\s*[:.]\s*([0-5]\d)(?![\d:.,])")


def _tijd_normaliseren(uur, minuut):
    uur, minuut = int(uur), int(minuut or 0)
    if 6 <= uur <= 23:
        return f"{uur:02d}:{minuut:02d}"
    return None


def vind_aanvangstijd(titelregels, regels):
    """Aanvangstijd als 'HH:MM'. In de titel mag een losse tijd ('10.00');
    in de rest van het document alleen met een duidelijk kenmerk ('aanvang',
    'om ... uur'), zodat bijbelverwijzingen als 'Johannes 3: 16' nooit als
    tijd tellen."""
    for regel in titelregels:
        if begint_met_boeknaam(regel):
            continue
        for patroon in (_TIJD_MET_CONTEXT, _TIJD_MET_UUR, _TIJD_LOS):
            m = patroon.search(regel)
            if m:
                t = _tijd_normaliseren(m.group(1), m.group(2))
                if t:
                    return t
    for regel in regels[:20]:
        if begint_met_boeknaam(regel) or parse_psalmregel(regel):
            continue
        for patroon in (_TIJD_MET_CONTEXT, _TIJD_MET_UUR):
            m = patroon.search(regel)
            if m:
                t = _tijd_normaliseren(m.group(1), m.group(2))
                if t:
                    return t
    return None


def vind_kerk(titelregels, regels):
    """'GK' (Grote Kerk) of 'MK' (Maranathakerk), eerst in de titel."""
    for bron in (titelregels, regels[:20]):
        for regel in bron:
            mk = re.search(r"maranatha|(?<![a-z])mk(?![a-z])", regel, re.IGNORECASE)
            gk = re.search(r"grote\s+kerk|(?<![a-z])gk(?![a-z])", regel, re.IGNORECASE)
            if mk and not gk:
                return "MK"
            if gk and not mk:
                return "GK"
    return None


def vind_organist(regels):
    for regel in regels:
        m = re.search(r"(?<![a-z])(?:organist(?:e)?|orgel(?:spel)?)\b\s*[:\-\u2013]?\s*(.+)", regel, re.IGNORECASE)
        if m:
            naam = _naam_na_titel(m.group(1))
            if naam:
                return naam
    return None


def bepaal_type_dienst(titelregels, regels, aanvangstijd=None):
    """'ochtend' of 'avond'. Voorrang: de aanvangstijd (voor 13:00 = ochtend,
    vanaf 13:00 = avond; de middagdienst wordt net als de avonddienst
    gehouden met de geloofsbelijdenis), dan woorden als 'zondagmorgen' in de
    titel, dan in het begin van de tekst. Woorden als 'Morgenzang' (Gezang 9),
    'Avondzang' (Gezang 12) en 'Avondmaal' tellen NIET mee."""
    if aanvangstijd:
        return "ochtend" if int(aanvangstijd[:2]) < 13 else "avond"
    ochtend = re.compile(r"(?<![a-z])(?:zondag)?(?:morgen|ochtend)(?:dienst)?(?![a-z])", re.IGNORECASE)
    avond = re.compile(r"(?<![a-z])(?:zondag)?avond(?:dienst)?(?![a-z])", re.IGNORECASE)
    for bron in (titelregels, regels[:12]):
        for regel in bron:
            o, a = ochtend.search(regel), avond.search(regel)
            if o and not a:
                return "ochtend"
            if a and not o:
                return "avond"
    return None


def _platte_regels(pad_of_tekst, is_bestandspad):
    """Alle niet-lege regels van het document als gewone tekst (zonder
    pandoc-opmaak); tabelrijen worden 'label : waarde'."""
    if is_bestandspad and pad_of_tekst.lower().endswith('.docx'):
        try:
            from docx import Document
            doc = Document(pad_of_tekst)
            regels = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for tabel in doc.tables:
                for rij in tabel.rows:
                    cellen = []
                    for c in rij.cells:
                        t = c.text.strip().replace("\n", " ")
                        if t and (not cellen or t != cellen[-1]):  # samengevoegde cellen niet dubbel
                            cellen.append(t)
                    if cellen:
                        regels.append(" : ".join(cellen))
            return regels
        except Exception:
            pass
        pad_of_tekst = docx_naar_tekst(pad_of_tekst)
    return [l.strip() for l in pad_of_tekst.splitlines() if l.strip()]


def parse_liturgie(pad_of_tekst, is_bestandspad=True):
    """pad_of_tekst: pad naar een .docx-bestand, OF platte tekst (zet dan
    is_bestandspad=False, bv. voor een geplakt e-mailtje)."""

    if is_bestandspad and pad_of_tekst.lower().endswith('.docx') and _docx_has_table(pad_of_tekst):
        result = parse_tabelvorm_docx(pad_of_tekst)
        result['_vorm'] = 'tabel'
        tekst_voor_extra_velden = docx_naar_tekst(pad_of_tekst)
    else:
        tekst_voor_extra_velden = docx_naar_tekst(pad_of_tekst) if (is_bestandspad and pad_of_tekst.lower().endswith('.docx')) else pad_of_tekst
        result = parse_losse_tekst(tekst_voor_extra_velden)
        result['_vorm'] = 'losse_tekst'

    # De voorganger halen we ALTIJD uit de titel van het document (nooit uit
    # de lopende tekst: daar staat "ds" ook in woorden als "Gods Woord" en
    # dan kwam er een stuk van het thema als predikant uit).
    voorganger = vind_voorganger_in_titel(pad_of_tekst, is_bestandspad)
    if voorganger:
        result['voorganger'] = voorganger

    # Dienstgegevens (datum, tijd, kerk, organist, type): eerst in de titel,
    # daarna in het begin van de tekst. Zo hoeft er op het controlescherm
    # niets meer ingetikt te worden dat al in de liturgie staat.
    titelregels = _titel_kandidaten(pad_of_tekst, is_bestandspad)
    regels = _platte_regels(pad_of_tekst, is_bestandspad)
    result['datum'] = vind_datum(titelregels) or vind_datum(regels[:15])
    result['aanvangstijd'] = vind_aanvangstijd(titelregels, regels)
    result['kerk'] = vind_kerk(titelregels, regels)
    result['organist'] = vind_organist(regels)
    result['type_dienst'] = result.get('type_dienst') or \
        bepaal_type_dienst(titelregels, regels, result['aanvangstijd'])

    # Catechismus-zondagen staan in avondmaalsliturgieën soms naast de
    # Bijbellezing. Het is een vaste verwijzing-dia, geen Bijbelhoofdstuk.
    for regel in tekst_voor_extra_velden.splitlines():
        m_cat = re.search(
            r"\b(?:heid(?:elbergse)?\.?\s*)?cat(?:echismus)?\.?\s*zondag\s*(\d{1,2})\b",
            regel, re.IGNORECASE)
        if m_cat:
            referentie = f"Heidelbergse Catechismus Zondag {m_cat.group(1)}"
            if referentie not in result.get("schriftlezingen", []):
                result.setdefault("schriftlezingen", []).append(referentie)

    # Bewaar alle liedmomenten in documentvolgorde. De standaarddienst blijft
    # de vaste secties gebruiken; andere dienstprofielen hebben soms meer
    # plekken en eigen momenten rond de avondmaalsviering.
    result['muziekmomenten'] = parse_muziekmomenten(tekst_voor_extra_velden)
    tekst_laag = tekst_voor_extra_velden.lower()
    if 'voortzetting' in tekst_laag and 'dankzegging' in tekst_laag:
        result['dienstprofiel'] = 'avondmaal_voortzetting'
    elif any(x in tekst_laag for x in ('heilig avondmaal', 'voorbereiden tafel', 'voorbereiding tafel')):
        result['dienstprofiel'] = 'avondmaal'
    else:
        result['dienstprofiel'] = 'standard'

    # thema en kerntekst verwisseld? Vangnet na het parsen (zie herstel_kern_thema)
    result.setdefault('parse_bevindingen', []).extend(herstel_kern_thema(result))

    return result


if __name__ == "__main__":
    import sys, json
    data = parse_liturgie(sys.argv[1])
    print(json.dumps(data, indent=2, ensure_ascii=False))
