"""
bijbeltekst.py
==============
Haalt de bijbelteksten voor de schriftlezingen uit de losse app "kerkdia-txt"
(zie kerkdia-txt/app/app.py in de aangeleverde 7z) en maakt er de dia-vulling
van: elke schriftlezing wordt een reeks dia's (een titeldia met de
verwijzing + tekstdia's met maximaal 8 regels per dia, versnummers in
superscript, precies zoals de app die al jaren netjes maakte).

Wat hier staat is een bewuste, 1-op-1 overname van de logica uit kerkdia-txt:
- De bijbelteksten komen uit JSON per boek/hoofdstuk in data/bijbeltekst/.
- De tekst wordt opgemaakt door woorden te meten met PIL op het lettertype
  Century Gothic Bold (data/fonts/GOTHICB.TTF) op 44 px. Die pixels zijn
  gekalibreerd op de echte dia-indeling (44 pt Century Gothic op een tekstvak
  van 10.124.387 EMU breed): max_breedte=775 px per regel, max 8 regels per
  dia -- de standaardwaarden van kerkdia-txt.
- Versnummers komen als losse superscript-runs (baseline 30000) in de tekst
  terecht, net als in de handmatig gemaakte dienstdia's.

Let op het auteursrecht: deze module vult dus wél echte bijbeltekst in. De
kerk gebruikt hiervoor haar eigen, legitiem verkregen HSV/NBV/NBV21-JSONs
(uit kerkdia-txt). Kies de vertaling per dienst.

GEBRUIK:
    from bijbeltekst import verwijzing_parsen, verwijzing_ophalen, \
        verwijzing_titeltekst, tekst_paginas

    v = verwijzing_parsen("Ezechiël 47: 1-12")     # -> dict of None
    verzen = verwijzing_ophalen("HSV", v)          # {'47': {'1': '...'}}
    paginas = tekst_paginas(verzen)                # regels per dia
"""
import json
import os
import re
import unicodedata
from functools import lru_cache

# ---------------------------------------------------------------------------
# Data-paden
# ---------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BIJBEL_DIR = os.path.join(BASE_DIR, "data", "bijbeltekst")
FONT_PAD = os.path.join(BASE_DIR, "data", "fonts", "GOTHICB.TTF")

# De vertalingen die in data/bijbeltekst/ staan; de mapnaam is de foldernaam
# uit kerkdia-txt (bible-HSV, bible-NBV, bible-NBV21).
VERTALINGEN = {"HSV": "bible-HSV", "NBV": "bible-NBV", "NBV21": "bible-NBV21"}
STANDAARD_VERTALING = "HSV"

# Extra veelgebruikte afkortingen die niet in books.json staan (kerkdienst-
# liturgieën zijn daar nogal vrij in): korte naam -> sleutel in books.json.
EXTRA_BOEK_ALIASEN = {
    "fil": "PHIL", "filipp": "PHIL",  # Filippenzen ('Fil.'/'Filipp.' komen niet in books.json voor)
    "matth": "MAT", "matt": "MAT", "mattheus": "MAT",  # Matteüs (ook oude spelling met -th-)
    "markus": "MARK", "mark": "MARK",  # het sjabloon/de liturgie schrijft Markus/Mark., books.json: Marcus
    "psalm": "PS",         # Psalmen (als schriftlezing gelezen)
    "openb": "REV",        # Openbaring
    "ezech": "EZEK",       # Ezechiël
    "efeze": "EPH",        # spellingvariant zonder ï (books.json: Efeziërs)
    "1 korinthe": "1CO", "2 korinthe": "2CO",
    "1 kon": "1KGS", "2 kon": "2KGS",
    "hooglied": "SNG",     # spellingvariant (books.json: Hooglied)
    # 'Petr.' / 'Petrus' zonder cijfer = 1 Petrus (in liturgieën staat vaak
    # alleen 'Petr. 1: 3-9' terwijl 1 Petrus bedoeld wordt)
    "petr": "1PE", "petrus": "1PE",
    # oude (NBG-)spellingen met -th- die nog veel in kerken leven
    "1 thessalonicenzen": "1TH", "2 thessalonicenzen": "2TH",
    "1 timotheus": "1TI", "2 timotheus": "2TI",
    "1 thess": "1TH", "2 thess": "2TH",
    "1 tess": "1TH", "2 tess": "2TH",
    # Spelling/afkortingen die in Nederlandse liturgieën veel voorkomen
    # (HSV schrijft 'Lukas', books.json heeft 'Lucas'; 'Mar.', 'Heb.', 'Jac.', ...)
    "lukas": "LUK", "luk": "LUK", "lk": "LUK",
    "mar": "MARK", "marc": "MARK", "mrk": "MARK", "mk": "MARK",
    "heb": "HEB", "jac": "JAS", "jak": "JAS", "tit": "TIT",
    "gen": "GEN", "ex": "EXO", "exod": "EXO", "lev": "LEV", "num": "NUM",
    "deut": "DEU", "joz": "JOS", "recht": "JDG",
    "1 sam": "1SA", "2 sam": "2SA", "1 kron": "1CH", "2 kron": "2CH",
    "neh": "NEH", "esther": "EST", "est": "EST",
    "spr": "PRO", "pred": "ECC", "hoogl": "SONG", "hgl": "SONG",
    "jes": "ISA", "jer": "JER", "klaagl": "LAM", "ez": "EZEK",
    "dan": "DAN", "hos": "HOS", "obad": "OBA", "ob": "OBA", "mich": "MIC",
    "nah": "NAH", "hab": "HAB", "zef": "ZEP", "zefanja": "ZEP", "sef": "ZEP",
    "hag": "HAG", "zach": "ZEC", "zac": "ZEC", "zacharia": "ZEC", "mal": "MAL",
    "mat": "MAT", "joh": "JOHN", "hand": "ACT", "rom": "ROM",
    "1 kor": "1CO", "2 kor": "2CO", "1 korinthiers": "1CO", "2 korinthiers": "2CO",
    "gal": "GAL", "ef": "EPH", "efez": "EPH", "kol": "COL",
    "1 tim": "1TI", "2 tim": "2TI", "filem": "PHLM", "filemon": "PHLM",
    "1 petr": "1PE", "2 petr": "2PE", "1 pet": "1PE", "2 pet": "2PE",
    "1 joh": "1JOHN", "2 joh": "2JOHN", "3 joh": "3JOHN", "jud": "JUD",
}


# ---------------------------------------------------------------------------
# Namen normaliseren (zonder accenten, kleine letters, puntjes weg)
# ---------------------------------------------------------------------------

def _norm(tekst):
    tekst = unicodedata.normalize("NFKD", tekst or "")
    tekst = "".join(c for c in tekst if not unicodedata.combining(c))
    tekst = tekst.replace("\u2013", "-").replace("\u2014", "-")  # – —
    tekst = tekst.replace(".", " ")
    return re.sub(r"\s+", " ", tekst.lower()).strip()


@lru_cache(maxsize=1)
def _boeken_laden():
    """Leest books.json + verses.json uit data/bijbeltekst/. Geeft een lijst
    met per boek: (sleutel, url-mapnaam, lange naam, genormaliseerde namen)."""
    try:
        met_books = json.load(open(os.path.join(BIJBEL_DIR, "books.json"), encoding="utf-8"))
    except FileNotFoundError:
        return [], {}
    boeken = []
    for key, info in met_books.items():
        kandidaten = {_norm(info["short"]), _norm(info["long"])}
        for alias, doel in EXTRA_BOEK_ALIASEN.items():
            if doel == key:
                kandidaten.add(_norm(alias))
        # langste eerst: '1 koningen' moet winnen van '1 kon'
        kandidaten = sorted((k for k in kandidaten if k), key=len, reverse=True)
        boeken.append({
            "key": key,
            "url": info.get("url", key),
            "long": info["long"],
            "namen": kandidaten,
        })
    # hoofdstukken-info (per boek: laatste hoofdstuknummer), niet cruciaal
    hoofdstukken = {}
    try:
        for entry in json.load(open(os.path.join(BIJBEL_DIR, "verses.json"), encoding="utf-8")):
            m = re.match(r"(\d+)\s*-\s*(\d+)", entry.get("verses", ""))
            if m:
                hoofdstukken[entry["book"]] = int(m.group(2))
    except (FileNotFoundError, KeyError):
        pass
    return boeken, hoofdstukken


def _lees_hoofdstuk(vertaling_map, boek_url, hoofdstuk):
    """Geeft de verzen van 1 hoofdstuk als lijst, of None als het bestand
    (of de vertaling) niet aanwezig is."""
    pad = os.path.join(BIJBEL_DIR, vertaling_map, boek_url, f"{hoofdstuk}.json")
    try:
        with open(pad, encoding="utf-8") as f:
            verzen = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None
    return [str(v).strip("\n ") for v in verzen]


# ---------------------------------------------------------------------------
# Verwijzing parsen: 'Ezechiël 47: 1-12', 'Fil. 1: 3 t/m 6', 'Psalmen 148: 1 - 14'
# ---------------------------------------------------------------------------

_CIJFER = r"\d{1,3}"
_SCHEIDING = r"(?:-|–|—|t/?m|tot\s+en\s+met)"
_MULTI_HFDST = re.compile(
    rf"^\s*({_CIJFER})\s*:\s*({_CIJFER})\s*{_SCHEIDING}\s*({_CIJFER})\s*:\s*({_CIJFER})")
_RANGE = re.compile(
    rf"^\s*({_CIJFER})\s*:\s*({_CIJFER})\s*(?:{_SCHEIDING}\s*({_CIJFER}))?")
_ALLES = re.compile(rf"^\s*({_CIJFER})\s*$")


_VERTALING_ACHTERAAN = re.compile(r"\s+(?:hsv|nbv21|nbv|nbg51|nbg|sv|bgt|niv)\s*$", re.IGNORECASE)


def _voorbewerk(tekst):
    """Maakt veelvoorkomende schrijfwijzen gelijk aan 'Boek H: V-V' voordat er
    geparsed wordt: '(HSV)'/'NBV' achteraan, opmerkingen tussen haakjes,
    'Johannes 3 vers 16 t/m 21', 'Joh. 3,16-21' (komma i.p.v. dubbele punt)."""
    t = (tekst or "").strip()
    t = re.sub(r"\s*\([^)]*\)", "", t)              # (de verloren zoon) / (HSV)
    t = _VERTALING_ACHTERAAN.sub("", t)
    t = t.replace("\u2013", "-").replace("\u2014", "-")
    t = re.sub(r"\b(?:hoofdstuk|hfdst|hfst)\.?\s*(?=\d)", "", t, flags=re.IGNORECASE)
    # 'vers 16' / 'vs. 16' / 'verzen 16-21' -> ': 16'
    t = re.sub(r"(\d)\s*(?:vers|verzen|vs|vv?)\.?\s+(?=\d)", r"\1: ", t, flags=re.IGNORECASE)
    # Nederlandse komma-notatie 'Johannes 3,16-21' (alleen als er nog geen ':' in staat)
    if ":" not in t:
        m = re.match(r"^(.*?[A-Za-z\u00c0-\u017f.])\s*(\d{1,3})\s*,\s*(\d{1,3}(?:\s*(?:-|t/?m)\s*\d{1,3})?)\s*$",
                     t, flags=re.IGNORECASE)
        if m:
            t = f"{m.group(1)} {m.group(2)}: {m.group(3)}"
    return t.strip()


def begint_met_boeknaam(tekst):
    """True als de tekst met een bekende boeknaam/afkorting begint ('Marcus 4: 35')."""
    norm = _norm(_voorbewerk(tekst))
    boeken, _ = _boeken_laden()
    for boek in boeken:
        for naam in boek["namen"]:
            if norm.startswith(naam + " ") and norm[len(naam) + 1:][:1].isdigit():
                return True
    return False


def splits_verwijzingen(tekst):
    """'Genesis 1: 1-5 en Marcus 4: 35-41' -> ['Genesis 1: 1-5', 'Marcus 4: 35-41'].
    Alleen splitsen waar erna weer een boeknaam volgt: 'Romeinen 3: 1 en 4' en
    'Johannes 3: 1-8, 14-16' blijven dus 1 verwijzing."""
    delen = re.split(r"(\s*(?:;|&|\+|\ben\b|,)\s*)", tekst or "", flags=re.IGNORECASE)
    uit = [delen[0]]
    for i in range(1, len(delen), 2):
        sep, seg = delen[i], delen[i + 1] if i + 1 < len(delen) else ""
        if seg.strip() and begint_met_boeknaam(seg):
            uit.append(seg)
        elif seg.strip() and re.match(r"^\d{1,3}\s*[:.]\s*\d", seg.strip()):
            # '1 Petr. 2:11-25 en 3:8-12': een nieuw hoofdstuk zonder boeknaam erbij is een
            # tweede stuk uit hetzelfde boek. Los van elkaar kunnen ze wel automatisch
            # gevuld worden, samen ('deels') niet.
            boek = re.match(r"^\s*((?:[1-3]\s*)?[A-Za-z\u00c0-\u017f][A-Za-z\u00c0-\u017f.]*)\s*\d", uit[-1])
            if boek and begint_met_boeknaam(uit[-1]):
                uit.append(boek.group(1).strip() + " " + seg.strip())
            else:
                uit[-1] += sep + seg
        else:
            uit[-1] += sep + seg
    return [u.strip() for u in uit if u.strip()]


def verwijzing_parsen(tekst):
    """Zet een bijbelverwijzing om naar een dict, of None als er geen
    (ondubbelzinnige) verwijzing in staat. Het dict bevat:
        boek_key, boek_url, boek_naam (lange naam), hoofdstuk_van, hoofdstuk_tm,
        vers_van, vers_tm, heel (True als het alleen hele verzen zijn),
        ruwe_tekst
    'heel' is False bij verwijzingen als 'Ezechiël 47: 9b+12' -- die kunnen
    niet automatisch gevuld worden (de JSON bevat alleen hele verzen)."""
    if not tekst:
        return None
    norm = _norm(_voorbewerk(tekst))
    boeken, _ = _boeken_laden()
    for boek in boeken:
        for naam in boek["namen"]:
            if norm == naam:
                rest = ""
                break
            if norm.startswith(naam + " ") and norm[len(naam) + 1:][:1].isdigit():
                rest = norm[len(naam):].strip()
                break
        else:
            continue
        break
    else:
        return None

    # 'Johannes 3: 16 e.v.' / '3: 16vv' -- einde onbekend: nooit stilletjes
    # tot 1 vers inkorten, dus niet automatisch invullen
    if re.search(r"\d\s*(?:e ?v|vv|ev)\s*$", rest):
        return None

    m = _MULTI_HFDST.match(rest)
    if m:
        hf, vf, ht, vt = (int(g) for g in m.groups())
        heel = True
    else:
        m = _RANGE.match(rest)
        if m:
            hf, vf = int(m.group(1)), int(m.group(2))
            ht = hf
            vt = int(m.group(3)) if m.group(3) else vf
        else:
            m = _ALLES.match(rest)
            if not m:
                return None
            hf = int(m.group(1))
            ht = hf
            vf = vt = None  # heel hoofdstuk; vers-tm wordt bepaald bij het ophalen
        # deels verzen ('Ezechiël 47: 9b+12'), losse opsommingen
        # ('3: 1 en 4') kunnen we niet automatisch vullen; 't/m' en
        # streepjes zijn gewone bereiken en dus wél heel
        heel = not re.search(r"\d[a-zA-Z]|[+,]|(?:^|\s)en(?:\s|$)", rest)
        # blijven er na het bereik nog cijfers over ('3: 1-8 14-16'), dan is
        # het geen enkel bereik meer
        if heel and re.search(r"\d", rest[m.end():]):
            heel = False
        # Een heel hoofdstuk ('Jesaja 53', 'Psalm 23') is gewoon een bereik van
        # vers 1 t/m het laatste vers; verwijzing_ophalen bepaalt het einde.

    if ht < hf or (ht == hf and vt is not None and vt < vf):
        return None

    # Deels-verwijzingen ('Ezechiël 47: 9b+12', 'Romeinen 3: 1 en 4') worden
    # bewaard als een lijst (versnummer, eventueel a/b-deel) zodat ze later
    # met met_delen=True wél opgehaald kunnen worden (zie verwijzing_ophalen).
    # 'heel' blijft dan False, zodat de gewone schriftlezing-paden (die geen
    # losse a/b-delen kunnen tonen) ze net als voorheen als placeholder laten.
    delen = None
    if not heel and hf == ht and vf is not None:
        m_verzen = re.match(rf"^\s*{hf}\s*:\s*(.+)$", rest)
        if m_verzen:
            lijst = re.sub(r"\s+en\s+", "+", m_verzen.group(1), flags=re.IGNORECASE)
            delen = []
            for token in re.split(r"\s*[,+]\s*", lijst.strip()):
                m_ber = re.match(r"^(\d{1,3})-(\d{1,3})$", token)
                if m_ber:  # bereik binnen de opsomming, bv. '3: 1-3 en 5'
                    a, b = int(m_ber.group(1)), int(m_ber.group(2))
                    if b < a:
                        delen = None
                        break
                    delen += [(nr, "") for nr in range(a, b + 1)]
                    continue
                m_deel = re.match(r"^(\d{1,3})([a-z])?$", token, re.IGNORECASE)
                if not m_deel:
                    delen = None
                    break
                delen.append((int(m_deel.group(1)), (m_deel.group(2) or "").lower()))
            if not delen:
                delen = None

    return {
        "boek_key": boek["key"],
        "boek_url": boek["url"],
        "boek_naam": boek["long"],
        "hoofdstuk_van": hf,
        "hoofdstuk_tm": ht,
        "vers_van": vf,
        "vers_tm": vt,
        "heel": heel,
        "delen": delen,
        "ruwe_tekst": tekst.strip(),
    }


def verwijzing_titeltekst(v):
    """Titel voor de titeldia van een schriftlezing, in dezelfde stijl als
    kerkdia-txt: 'Ezechiël 47 : 1 - 12', 'Psalmen 148 : 1 - 14'."""
    if v is None:
        return ""
    if v["vers_van"] is None:  # heel hoofdstuk
        t = f"{v['boek_naam']} {v['hoofdstuk_van']}"
        return t + (f"  t/m  {v['hoofdstuk_tm']}" if v["hoofdstuk_tm"] != v["hoofdstuk_van"] else "")
    t = f"{v['boek_naam']} {v['hoofdstuk_van']} : {v['vers_van']}"
    if v["hoofdstuk_tm"] != v["hoofdstuk_van"]:
        t += f"  t/m  {v['hoofdstuk_tm']} : "
    else:
        t += " - "
    t += str(v["vers_tm"])
    return t


def verwijzing_ophalen(vertaling, v, met_delen=False):
    """Haalt de verzen op als {'<hoofdstuk>': {'<vers>': 'tekst'}}, in de
    zelfde vorm als kerkdia-txt. Geeft None als de vertaling/boeken niet
    aanwezig zijn of de verwijzing buiten de hoofdstukken valt.

    Met met_delen=True kunnen ook deels-verwijzingen opgehaald worden
    ('Ezechiël 47: 9b+12'): de genoemde hele verzen plus de a/b-delen, die
    dan op de zinsgrens dichtst bij het midden van het vers gesplitst
    worden (best effort -- zie _vers_deel)."""
    vertaling_map = VERTALINGEN.get(vertaling or "")
    if vertaling_map is None or v is None:
        return None
    if not v["heel"]:
        # alleen deels-verwijzingen binnen één hoofdstuk en expliciet gevraagd
        if not (met_delen and v.get("delen") and v["hoofdstuk_van"] == v["hoofdstuk_tm"]):
            return None
        hfdst = v["hoofdstuk_van"]
        hoofdstuk_verzen = _lees_hoofdstuk(vertaling_map, v["boek_url"], hfdst)
        if hoofdstuk_verzen is None:
            return None
        deel = {}
        for vers, letter in v["delen"]:
            if vers - 1 >= len(hoofdstuk_verzen):
                return None
            tekst = hoofdstuk_verzen[vers - 1].strip()
            if letter:
                tekst = _vers_deel(tekst, letter)
                if not tekst:
                    return None
            deel[f"{vers}{letter}"] = tekst
        return {str(hfdst): deel}
    verzen = {}
    for hfdst in range(v["hoofdstuk_van"], v["hoofdstuk_tm"] + 1):
        hoofdstuk_verzen = _lees_hoofdstuk(vertaling_map, v["boek_url"], hfdst)
        if hoofdstuk_verzen is None:
            return None
        if v["vers_van"] is None:  # heel hoofdstuk
            start, eind = 1, len(hoofdstuk_verzen)
        elif hfdst == v["hoofdstuk_van"]:
            start, eind = v["vers_van"], (v["vers_tm"] if hfdst == v["hoofdstuk_tm"] else len(hoofdstuk_verzen))
        elif hfdst == v["hoofdstuk_tm"]:
            start, eind = 1, v["vers_tm"]
        else:
            start, eind = 1, len(hoofdstuk_verzen)
        if start > len(hoofdstuk_verzen):
            return None
        deel = {}
        for nr in range(start, eind + 1):
            if nr - 1 >= len(hoofdstuk_verzen):
                return None
            deel[str(nr)] = hoofdstuk_verzen[nr - 1]
        verzen[str(hfdst)] = deel
    return verzen


def _vers_deel(tekst, letter):
    """Splits een (hele) vers-tekst in een a- of b-deel. De grens ligt op de
    zinsgrens die het dichtst bij het midden van het vers zit; een vers dat
    uit maar één zin bestaat wordt rond de helft op een woordgrens geknipt.
    Dit is best effort: de kerk kiest zulke delen zelf ook vaak anders."""
    tekst = (tekst or "").strip()
    if not tekst:
        return None
    zinnen = re.split(r"(?<=[.;!?])\s+", tekst)
    helft = len(tekst) // 2
    if len(zinnen) < 2:
        # één lange zin: knippen op de eerstvolgende (of anders laatste)
        # woordgrens rond de helft
        grens = None
        for pos in range(helft, len(tekst)):
            if tekst[pos] == " ":
                grens = pos
                break
        if grens is None:
            for pos in range(helft, 0, -1):
                if tekst[pos] == " ":
                    grens = pos
                    break
        if grens is None:  # niet te splitsen (1 woord)
            return tekst if letter == "b" else None
        return (tekst[grens:].strip() if letter == "b" else tekst[:grens].strip()) or None
    # meerdere zinnen: kies de zinsgrens het dichtst bij het midden
    haltes = []
    cum = 0
    for zin in zinnen[:-1]:
        cum += len(zin) + 1  # + spatie
        haltes.append(cum)
    grens = min(range(len(haltes)), key=lambda i: abs(haltes[i] - helft))
    grens = haltes[grens]
    return (tekst[grens:].strip() if letter == "b" else tekst[:grens].strip()) or None


# ---------------------------------------------------------------------------
# Opmaak: woorden meten + regels/dia's verdelen (port van kerkdia-txt)
# ---------------------------------------------------------------------------

_MAX_BREEDTE = 775   # px, gekalibreerd op 44 pt Century Gothic Bold
_MAX_REGELS = 8      # regels per dia


@lru_cache(maxsize=1)
def _meetfont():
    from PIL import ImageFont
    return ImageFont.truetype(FONT_PAD, 44)


def tekst_paginas(verzen, max_breedte=_MAX_BREEDTE, max_regels=_MAX_REGELS):
    """Verdeelt de verzen over dia's. Geeft een lijst pagina's terug; elke
    pagina is een lijst regels, elke regel is een lijst runs van
    (soort, tekst) met soort 'n' (normaal, vet) of 'v' (versnummer in
    superscript). Dit is een getrouwe port van de tekst-lus uit kerkdia-txt
    (generate_presentation), met dezelfde structuur (buitenste lus = één
    regel of één nieuwe dia per keer), zodat de uitkomst gelijk is aan de
    dia's die de kerk al jaren met die app maakt.

    De 'verzen' moeten de vorm {'<hfdst>': {'<vers>': 'tekst'}} hebben
    (zie verwijzing_ophalen)."""
    if not verzen:
        return []
    font = _meetfont()

    hfdst_van = min(int(h) for h in verzen)
    hfdst_tm = max(int(h) for h in verzen)
    hfdst = hfdst_van
    # let op: numeriek vergelijken -- min() over de string-sleutels zou
    # '10' < '6' geven en verzen stilletjes overslaan
    vers = int(min(int(k) for k in verzen[str(hfdst_van)])) - 1

    paginas = []      # afgesloten dia's, elk een lijst regels
    cur_pagina = None
    regel = None
    regel_breedte = 0
    run_open = False  # huidige regel heeft al tekst (analoog aan 'run' in kerkdia-txt)
    regels = 1
    vers_tekst = []
    volgende_woord = ""
    nieuwe_zin = True
    nummer_toevoegen = False
    klaar = False

    while True:  # buitenste lus: 1x per regel, of nieuwe dia als die vol is
        if cur_pagina is None:
            cur_pagina = []
            paginas.append(cur_pagina)
        if regel is None:
            regel = []
            regel_breedte = 0
        run_open = False

        while True:  # binnenste lus: woorden op deze regel
            if volgende_woord == "":
                if len(vers_tekst) == 0:
                    if str(vers + 1) in verzen[str(hfdst)]:
                        vers += 1
                    elif hfdst_tm > hfdst and str(hfdst + 1) in verzen:
                        # volgende hoofdstuk: verzen tellen daar opnieuw vanaf
                        # het eerste vers van dat hoofdstuk
                        hfdst += 1
                        vers = int(min(int(k) for k in verzen[str(hfdst)]))
                    else:
                        klaar = True
                        break
                    vers_tekst = verzen[str(hfdst)][str(vers)].split()
                    nummer_toevoegen = True
                    nieuwe_zin = True
                volgende_woord = vers_tekst.pop(0)

            # voorkom een eenzaam eerste woord van een nieuwe zin op de
            # laatste regel van een dia
            if regels == max_regels and nieuwe_zin and len(vers_tekst) > 1:
                volgende_woord += " " + vers_tekst.pop(0)

            # dezelfde ruimte-logica als kerkdia-txt: alleen een woord aan
            # het begin van een regel (zonder versnummer) verliest z'n
            # aanloopspatie
            if run_open or nummer_toevoegen:
                volgende_woord = " " + volgende_woord.lstrip()
            else:
                volgende_woord = volgende_woord.lstrip()

            nieuwe_zin = volgende_woord[-1:] in ".,"

            meet_woord = volgende_woord
            if nummer_toevoegen:
                # breedte van het superscript-versnummer meerekenen
                meet_woord = (" (" + meet_woord) if vers < 10 else (" ((" + meet_woord)

            regel_breedte += font.getlength(meet_woord)

            if regel_breedte > max_breedte and run_open:
                # deze regel is vol: afsluiten en (via de buitenste lus)
                # verder op een nieuwe regel -- of nieuwe dia
                cur_pagina.append(regel)
                if regels >= max_regels:
                    cur_pagina = None
                    regels = 1
                else:
                    regels += 1
                regel = None
                break

            if nummer_toevoegen:
                nummer_toevoegen = False
                regel.append(("n", " "))
                regel.append(("v", str(vers)))
                run_open = False

            # woord toevoegen (verleng het lopende run, net als kerkdia-txt
            # dat met run.text doet)
            if run_open and regel and regel[-1][0] == "n":
                vorige = regel[-1][1]
                regel[-1] = ("n", vorige + volgende_woord)
            else:
                regel.append(("n", volgende_woord))
                run_open = True
            volgende_woord = ""

        if klaar:
            break

    # laatste (nog openstaande) regel en dia afsluiten
    if regel:
        cur_pagina.append(regel)
    return paginas


def regels_naar_xml(pagina):
    """Zet 1 pagina (regels met runs) om naar de <a:p>-XML die in een dia met
    de ctrTitle-placeholder (44 pt Century Gothic, zoals in het kerksjabloon)
    geplakt kan worden. Zelfde run-structuur als de bestaande
    schriftlezingdia's: tekst vet, versnummers als superscript (baseline)."""
    from xml.sax.saxutils import escape
    stukken = []
    for regel in pagina:
        stukken.append("<a:p>")
        for soort, tekst in regel:
            if not tekst:
                continue
            if soort == "v":
                rpr = ' b="0" i="1" baseline="30000"'
            else:
                rpr = ' b="1"'
            stukken.append(f'<a:r><a:rPr{rpr}/><a:t>{escape(tekst)}</a:t></a:r>')
        stukken.append("</a:p>")
    return "".join(stukken)
