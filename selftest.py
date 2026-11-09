"""
selftest.py
===========
Snelle zelfcontrole van de hele app. Draai hem na elke wijziging en na het bijwerken van
data/ (nieuw jaarrooster, ander sjabloon):

    python selftest.py

Controleert:
  1. Rooster: elke voorzang in data/psalmrooster_*.json is te lezen door de parser en
     staat (voor psalmen/gezangen uit de eigen bibliotheek) ook echt in de index.
  2. Herkenning: kerntekst / thema / schriftlezing worden goed uit elkaar gehouden
     (tabelvorm en losse tekst, met allerlei labels).
  3. Sjabloon: staan de dia-nummers uit template_map.json nog op de juiste dia's?
  4. Data: lopen de jaarroosters nog lang genoeg door?
Geeft exitcode 1 bij een echte fout (waarschuwingen tellen niet mee).
"""
import glob
import json
import os
import re
import sys
import tempfile
import zipfile
from datetime import date, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
DATA = os.path.join(BASE, "data")

fouten, waarschuwingen = [], []


def fout(tekst):
    fouten.append(tekst)
    print("  FOUT  ", tekst)


def let_op(tekst):
    waarschuwingen.append(tekst)
    print("  LET OP", tekst)


# ---------------------------------------------------------------- 1. rooster
def test_rooster():
    print("\n1. Psalmrooster tegen de bibliotheek")
    from assemble import _psalm_entry
    from liturgie_parser import parse_psalmregel
    psalm_index = json.load(open(os.path.join(DATA, "psalm_index.json"), encoding="utf-8"))
    gpad = os.path.join(DATA, "gezang_index.json")
    gezang_index = json.load(open(gpad, encoding="utf-8")) if os.path.exists(gpad) else {}
    totaal = weerklank = ok = 0
    for pad in sorted(glob.glob(os.path.join(DATA, "psalmrooster_*.json"))):
        for r in json.load(open(pad, encoding="utf-8")):
            v = r.get("voorzang")
            if not v:
                continue
            totaal += 1
            p = parse_psalmregel(v)
            if not p:
                let_op(f"{r['datum']} {r['tijd']} {r['kerk']}: voorzang '{v}' is geen psalm/gezang "
                       "(hoort de app te melden, niet automatisch te vullen)")
                continue
            if p["bundel"] == "Wk":
                weerklank += 1
                continue
            index = gezang_index if p["soort"] == "Gezang" else psalm_index
            ontbreekt = []
            for vn in p["verzen"]:
                e = _psalm_entry(index, p["nummer"], [vn])
                if e is None or vn not in {x["vers"] for x in e["verses"]}:
                    ontbreekt.append(vn)
            if ontbreekt:
                let_op(f"{r['datum']} {r['tijd']} {r['kerk']}: '{v}' -- vers {ontbreekt} staat niet in de "
                       "bibliotheek (dia moet zelf aangeleverd worden)")
            else:
                ok += 1
    print(f"  {totaal} voorzangen: {ok} volledig automatisch, {weerklank} Weerklank (zipje nodig)")


# ------------------------------------------------------------- 2. herkenning
def _tabel(rijen, titel="Liturgie zondag 27 sept - ds. J. Jansen"):
    from docx import Document
    from liturgie_parser import parse_liturgie
    d = Document()
    d.add_paragraph(titel)
    t = d.add_table(rows=0, cols=2)
    for a, b in rijen:
        r = t.add_row()
        r.cells[0].text, r.cells[1].text = a, b
    with tempfile.TemporaryDirectory() as tmp:
        p = os.path.join(tmp, "t.docx")
        d.save(p)
        return parse_liturgie(p)


def _los(regels, titel="Liturgie zondag 27 sept - ds. J. Jansen"):
    from docx import Document
    from liturgie_parser import parse_liturgie
    d = Document()
    d.add_paragraph(titel)
    for r in ["Psalm 100: 1"] + regels:
        d.add_paragraph(r)
    with tempfile.TemporaryDirectory() as tmp:
        p = os.path.join(tmp, "l.docx")
        d.save(p)
        return parse_liturgie(p)


def test_herkenning():
    print("\n2. Kerntekst / thema / schriftlezing uit elkaar houden")
    cases = [
        ("tabel: Tekstkeuze = verwijzing", _tabel([("Tekstkeuze", "Johannes 3: 16")]), dict(kerntekst="Johannes 3: 16", thema=None)),
        ("tabel: Tekstkeuze met 'Thema:'", _tabel([("Tekstkeuze", "Thema: 'Vrede'")]), dict(thema="Vrede", kerntekst=None)),
        ("tabel: Preektekst zonder verwijzing = thema", _tabel([("Preektekst", "Vrede in de storm")]), dict(thema="Vrede in de storm", kerntekst=None)),
        ("tabel: Thema is een verwijzing = kerntekst", _tabel([("Thema", "Efeziers 2: 8-10")]), dict(kerntekst="Efeziers 2: 8-10", thema=None)),
        ("tabel: Bijbelgedeelte + kerntekst vers 16", _tabel([("Bijbelgedeelte", "Johannes 3: 1-21 (kerntekst: vers 16)")]), dict(schriftlezingen=["Johannes 3: 1-21"], kerntekst="Johannes 3: 16")),
        ("tabel: Schriftlezing en tekst", _tabel([("Schriftlezing en tekst", "Marcus 4: 35-41")]), dict(schriftlezingen=["Marcus 4: 35-41"], kerntekst="Marcus 4: 35-41")),
        ("tabel: Kerntekst met citaat", _tabel([("Kerntekst", "'Uw woord is een lamp' (Psalm 119: 105)")]), dict(kerntekst="Psalm 119: 105")),
        ("tabel: lezing + kerntekst + thema los", _tabel([("Schriftlezing", "Genesis 1: 1-5"), ("Kerntekst", "Johannes 1: 14"), ("Thema", "Genade")]), dict(schriftlezingen=["Genesis 1: 1-5"], kerntekst="Johannes 1: 14", thema="Genade")),
        ("tabel: Thema met verwijzing erachter", _tabel([("Thema", "Liefde (Johannes 3: 16)")]), dict(thema="Liefde", kerntekst="Johannes 3: 16")),
        ("tabel: Bijbeltekst, 1 vers = kerntekst", _tabel([("Bijbeltekst", "Johannes 3: 16")]), dict(kerntekst="Johannes 3: 16", schriftlezingen=[])),
        ("tabel: Bijbeltekst, bereik = lezing", _tabel([("Bijbeltekst", "Genesis 1: 1-5")]), dict(kerntekst=None, schriftlezingen=["Genesis 1: 1-5"])),
        ("tabel: label Tekst + verwijzing", _tabel([("Tekst", "Romeinen 8: 28")]), dict(kerntekst="Romeinen 8: 28")),
        ("tabel: label Tekst zonder verwijzing", _tabel([("Tekst", "Zondag 1")]), dict(kerntekst=None, thema=None)),
        ("tabel: Kerntekst zonder verwijzing = thema", _tabel([("Kerntekst", "Genade & waarheid")]), dict(thema="Genade & waarheid", kerntekst=None)),
        ("tabel: thema en tekst allebei ingevuld", _tabel([("Thema", "Liefde"), ("Kerntekst", "Alleen door genade")]), dict(thema="Liefde", kerntekst="Alleen door genade")),
        ("los: Preek: zonder verwijzing = thema", _los(["Preek: Genade & waarheid"]), dict(thema="Genade & waarheid", kerntekst=None)),
        ("los: Kerntekst + Thema", _los(["Kerntekst: Johannes 1: 14", "Thema: 'Genade'"]), dict(kerntekst="Johannes 1: 14", thema="Genade")),
        ("los: Thema: verwijzing = kerntekst", _los(["Thema: Johannes 3: 16"]), dict(kerntekst="Johannes 3: 16", thema=None)),
        ("los: Schriftlezing, kerntekst vers 16", _los(["Schriftlezing: Johannes 3: 1-21, kerntekst vers 16"]), dict(schriftlezingen=["Johannes 3: 1-21"], kerntekst="Johannes 3: 16")),
        ("los: Lezing ... en tekst: ...", _los(["Lezing: Genesis 1: 1-5 en tekst: Marcus 4: 35-41"]), dict(schriftlezingen=["Genesis 1: 1-5"], kerntekst="Marcus 4: 35-41")),
        ("los: Prediking: verwijzing", _los(["Prediking: Johannes 3: 16"]), dict(kerntekst="Johannes 3: 16")),
        ("los: Tekst: verwijzing", _los(["Tekst: Johannes 3: 16"]), dict(kerntekst="Johannes 3: 16", schriftlezingen=[])),
        ("los: Onderwerp met verwijzing", _los(["Onderwerp: 'Vrede' - Filippenzen 4: 7"]), dict(thema="Vrede", kerntekst="Filippenzen 4: 7")),
        ("los: Schriftlezing en tekst", _los(["Schriftlezing en tekst: Marcus 4: 35-41"]), dict(schriftlezingen=["Marcus 4: 35-41"], kerntekst="Marcus 4: 35-41")),
        ("los: Thema=verwijzing + Preek=tekst", _los(["Thema: Johannes 3: 16", "Preek: Genade & waarheid"]), dict(kerntekst="Johannes 3: 16", thema="Genade & waarheid")),
        # boekloze kerntekst ("Tekst 2:13a '...'") krijgt het boek van de schriftlezing; een lezing
        # 'A 2:11-25 en 3:8-12' wordt twee lezingen (los kunnen ze wel automatisch gevuld worden)
        ("los: tekst zonder boek op aparte regel", _los(["Schriftlezing: 1 Petr. 2:11-25 en 3:8-12", "Tekst 2:13a 'Onderwerp u dan omwille van de Heere'"]), dict(schriftlezingen=["1 Petr. 2:11-25", "1 Petr. 3:8-12"], kerntekst="1 Petrus 2: 13a", thema=None)),
        ("los: tekst zonder boek achter de lezing", _los(["Schriftlezing: 1 Petr. 2:11-25 en 3:8-12; Tekst 2:13a 'Onderwerp u dan'"]), dict(schriftlezingen=["1 Petr. 2:11-25", "1 Petr. 3:8-12"], kerntekst="1 Petrus 2: 13a", thema=None)),
        ("los: tekst VOOR de lezing", _los(["Tekst 3:8 'Wees eensgezind'", "Schriftlezing: 1 Petr. 2:11-25 en 3:8-12"]), dict(kerntekst="1 Petrus 3: 8", thema=None)),
        ("los: Kerntekst: vers 13", _los(["Schriftlezing: 1 Petr. 2:11-25", "Kerntekst: vers 13"]), dict(kerntekst="1 Petrus 2: 13")),
        ("los: lezing (tekst: 2:13a)", _los(["Lezing: 1 Petr. 2:11-25 (tekst: 2:13a)"]), dict(schriftlezingen=["1 Petr. 2:11-25"], kerntekst="1 Petrus 2: 13a")),
        ("los: boekloze tekst zonder lezing blijft kerntekst", _los(["Tekst 2:13a 'Onderwerp u dan'"]), dict(kerntekst="2:13a", thema=None, schriftlezingen=[])),
        ("los: 'thema' midden in een liedregel", _los(["Psalm 25: 1 (thema van de dag)", "Preek: Genade"]), dict(thema="Genade")),
    ]
    for naam, res, verwacht in cases:
        slecht = {k: (res.get(k), v) for k, v in verwacht.items() if res.get(k) != v}
        if slecht:
            fout(f"{naam}: {slecht}")
    print(f"  {len(cases) - len([1 for n, r, v in cases if any(r.get(k) != x for k, x in v.items())])}"
          f" van {len(cases)} herkenningsgevallen goed")

    # regels als 'Psalm 100: 1, 2 (zie ook Joh 3)' mogen de '3' niet als vers zien
    from liturgie_parser import parse_psalmregel
    for tekst, verwacht in [("WPs 135: 2", ("Wk", "Psalm", [2])), ("Gz 2: 3", ("OB", "Gezang", [3])),
                            ("WLd 297: 9", ("Wk", "Lied", [9])),
                            ("Psalm 100: 1, 2 (zie ook Joh 3)", ("OB", "Psalm", [1, 2])),
                            ("Ps. 100: 1 en 4her", ("OB", "Psalm", [1, 4]))]:
        p = parse_psalmregel(tekst)
        got = p and (p["bundel"], p["soort"], p["verzen"])
        if got != verwacht:
            fout(f"parse_psalmregel('{tekst}') gaf {got}, verwacht {verwacht}")


# ------------------------------------------------------------------ 3. sjabloon
def test_sjabloon():
    print("\n3. Sjabloon tegen template_map.json")
    tm = json.load(open(os.path.join(DATA, "template_map.json"), encoding="utf-8"))
    pptx = os.path.join(DATA, tm["template_file"])
    if not os.path.exists(pptx):
        fout(f"sjabloon {tm['template_file']} ontbreekt in data/")
        return
    z = zipfile.ZipFile(pptx)
    namen = set(z.namelist())

    def tekst(n):
        naam = f"ppt/slides/slide{n}.xml"
        if naam not in namen:
            return None
        return " ".join(re.findall(r"<a:t>([^<]*)</a:t>", z.read(naam).decode("utf-8")))

    verwacht = [
        ("titeldia", [tm["titel_dia"]["slide"]], ["Welkom in deze dienst", "Aanvang"]),
        ("welkomdia", [tm["welkom_dia"]["slide"]], ["Voorganger", "Organist"]),
        ("orde van dienst", tm["orde_van_dienst_dias"], ["Orde van dienst"]),
        ("wet", [tm["wet_en_geloofsbelijdenis"]["wet_slide"]], ["Wet"]),
        ("geloofsbelijdenis", [tm["wet_en_geloofsbelijdenis"]["geloofsbelijdenis_slide"]], ["Geloofsbelijdenis"]),
        ("collecte", tm["collecte_doel_dias"], ["Rondgang"]),
        ("schriftlezing-overzicht", [tm["schriftlezing"]["overzicht_dia"]], ["Schriftlezingen"]),
        ("kerntekst", tm["schriftlezing"]["kerntekst_dias"], ["Kerntekst"]),
        ("verkondiging", [tm["schriftlezing"]["verkondiging_referentie_dia"]], ["Verkondiging"]),
        ("thema", tm["schriftlezing"]["thema_dias"][1:], ["Thema"]),
    ]
    for naam, dias, woorden in verwacht:
        for n in dias:
            t = tekst(n)
            if t is None:
                fout(f"sjabloon: dia {n} ({naam}) bestaat niet meer -- template_map.json is niet meer bijgewerkt")
            else:
                for w in woorden:
                    if w.lower() not in t.lower():
                        fout(f"sjabloon: dia {n} moet '{naam}' zijn maar bevat niet '{w}' (dia's verschoven?)")
    for sectie, cfg in tm["secties"].items():
        if sectie.startswith("_"):
            continue
        for n in cfg["titel_dias"] + cfg["oude_versdias"]:
            if tekst(n) is None and f"ppt/slides/slide{n}.xml" not in namen:
                fout(f"sjabloon: dia {n} (sectie {sectie}) bestaat niet")
    # tekstruns met kenmerken (<a:t xml:space=...>) of leeg (<a:t/>) worden door de vervang-functies niet gezien
    for n in {n for _, dias, _ in verwacht for n in dias}:
        naam = f"ppt/slides/slide{n}.xml"
        if naam in namen:
            xml = z.read(naam).decode("utf-8")
            if re.search(r"<a:t\s[^>]*>|<a:t/>", xml):
                let_op(f"sjabloon: dia {n} heeft <a:t ...>/<a:t/>-runs; de vervang-functies zien die niet")
    print("  sjabloon gecontroleerd")


# ------------------------------------------------------------------- 4. data
def test_data():
    print("\n4. Jaarroosters")
    vandaag = date.today()
    grens = (vandaag + timedelta(days=60)).isoformat()
    for naam, sleutel in [("psalmrooster_*.json", "datum"), ("collecterooster_*.json", None)]:
        eind = None
        for pad in glob.glob(os.path.join(DATA, naam)):
            d = json.load(open(pad, encoding="utf-8"))
            datums = [r[sleutel] for r in d] if sleutel else list(d.keys())
            if datums:
                eind = max(datums + ([eind] if eind else []))
        if eind is None:
            fout(f"geen {naam} in data/")
        elif eind < vandaag.isoformat():
            fout(f"{naam}: loopt tot {eind}, dat is al voorbij")
        elif eind < grens:
            let_op(f"{naam}: loopt tot {eind} -- laad binnenkort het nieuwe jaarrooster in")
        else:
            print(f"  {naam}: loopt tot {eind}")


if __name__ == "__main__":
    for test in (test_rooster, test_herkenning, test_sjabloon, test_data):
        try:
            test()
        except Exception as e:  # een crash in een test is zelf ook een fout
            import traceback
            traceback.print_exc()
            fout(f"{test.__name__} crashte: {e}")
    print(f"\n{len(fouten)} fout(en), {len(waarschuwingen)} waarschuwing(en)")
    sys.exit(1 if fouten else 0)
