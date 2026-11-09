"""
gezang_indexer.py
=================
Bouwt (net als psalm_indexer.py voor de psalmen) een index van "welk vers staat
op welke dia" voor de Gezangen uit de eigen bibliotheek (Gezang 1 t/m 12,
Oude Berijming -- "Gezang 5, Gebed des Heeren.pptx" enz.).

Waarom een APARTE indexer en niet gewoon psalm_indexer.py?
  Gezangen zijn onregelmatiger dan psalmen. Een vers telt niet altijd evenveel
  dia's (Gezang 3: 3-3-3-3-1 dia's, Gezang 6: 4-2), sommige gezangen hebben
  maar 1 vers dat over meerdere dia's loopt (Gezang 8, 10, 11), en in Gezang 4
  slaat de bron zelf een versnummer over (dia met "1." en dia met "3."). De
  psalm-aanpak ("elk vers heeft evenveel dia's") geeft daar dus fout resultaat.
  Hier werken we daarom met START-DIA'S: een vers loopt van de dia waar het
  versnummer staat tot de dia vóór het volgende versnummer.

AANPAK:
1. Elke .pptx wordt gerenderd (LibreOffice + pdftoppm) en geOCR't -- dezelfde
   stappen als psalm_indexer.py (die functies worden hergebruikt).
2. Per dia zoeken we in de eerste OCR-regels naar een versnummer aan het begin
   van de eerste tekstregel.
3. Uit de betrouwbare detecties (strikt oplopend) worden de start-dia's van de
   verzen bepaald. Ontbrekende detecties worden alleen aangevuld als de
   afstand tussen twee opeenvolgende, gewoon-doorlopende verzen (1->2->3) een
   vast aantal dia's is (en nooit als er een versnummer wordt overgeslagen).
4. Alles wat niet zeker is (nummer-sprong, geen enkele detectie, veel dia's
   maar maar één herkend vers, ...) komt in review_needed_gezang.json.
5. gezang_correcties.json (optioneel, in dezelfde output-map) bevat handmatig
   gecontroleerde indelingen en wint altijd van de OCR-uitkomst.

GEBRUIK:
    python gezang_indexer.py /pad/naar/gezangen_map data/

    gezangen_map = map met de .pptx-bestanden (mag ook submappen hebben).
    Output in de output-map (bv. data/):
      gezang_index.json, review_needed_gezang.json,
      render_cache/Gezang_<n>__<titel>/slide-NN.jpg

Het bestandsnummer ("Gezang 8, ...") bepaalt het gezangnummer -- NIET de tekst
op de titeldia (in "Gezang 8, Bedezang voor de predicatie.pptx" staat op de
titeldia per ongeluk "Gezang 3").
"""
import glob
import json
import os
import re
import shutil
import sys

from psalm_indexer import convert_pptx_to_slides, ocr_lines

GEZANG_NAAM = re.compile(r"gezang\s*(\d{1,3})", re.IGNORECASE)

# Versnummer aan het begin van de eerste tekstregel: "1.", "2,", "3 " -- gevolgd
# door echte tekst (minstens 2 letters, eventueel na 1 los OCR-tekentje, want
# "1 0 God" is hoe OCR "1. O God" soms leest).
_START_PATROON = re.compile(r"^\s*(\d{1,2})\s*[.,]?\s+(?:\S{1,2}\s+)?['\u2018\u2019A-Za-z]{2,}")
_OCR_REGELS_DOORZOEKEN = 4  # het versnummer staat in de eerste regels; verder zoeken geeft alleen ruis


def gezang_nummer(bestandsnaam):
    m = GEZANG_NAAM.search(bestandsnaam)
    return int(m.group(1)) if m else None


def cache_mapnaam(bestandsnaam_zonder_ext):
    """Zelfde sanitizing als assemble._verse_image_paths als eerste probeert."""
    return re.sub(r"[^\w]", "_", bestandsnaam_zonder_ext)


def detecteer_start(lines):
    """Geeft het versnummer als deze dia met een versnummer begint, anders None."""
    for line in lines[:_OCR_REGELS_DOORZOEKEN]:
        m = _START_PATROON.match(line)
        if m and int(m.group(1)) >= 1:
            return int(m.group(1))
    return None


def _oplopende_keten(detecties):
    """Houdt alleen detecties over die strikt oplopen in zowel dia als vers
    (OCR-ruis als '0' of een nummer dat terugspringt valt hierdoor af)."""
    keten = []
    laatste_nummer = 0
    for dia, nr in sorted(set(detecties)):
        if nr > laatste_nummer:
            keten.append((dia, nr))
            laatste_nummer = nr
    return keten


def bouw_gezang_index(n_content_slides, detecties):
    """Van (dia, versnummer)-detecties naar {'verses': [{'vers', 'slides'}...]}.
    Dia 1 = titeldia; inhoud loopt van dia 2 t/m n_content_slides + 1."""
    laatste_dia = n_content_slides + 1
    keten = _oplopende_keten(detecties)
    redenen = []

    if not keten:
        # niets herkend: de conventie is dat vers 1 op dia 2 begint. Alles op 1
        # hoop = 1 vers. Bij meer dan 3 dia's is dat niet aannemelijk.
        starts = [(2, 1)]
        if n_content_slides > 3:
            redenen.append("geen enkel versnummer herkend")
        vertrouwen = "gemiddeld" if n_content_slides <= 3 else "laag"
        return _maak_resultaat(starts, laatste_dia, vertrouwen, redenen)

    starts = list(keten)
    if starts[0][0] != 2:
        if starts[0][1] == 1:
            redenen.append(f"vers 1 pas op dia {starts[0][0]} herkend (verwacht: dia 2)")
        else:
            # eerste herkende vers is niet vers 1 en niet op dia 2: verse(n)
            # ervoor zijn gemist -- dia 2 t/m start-1 als vers 1 opvatten
            starts.insert(0, (2, 1))
            redenen.append("vers 1 niet herkend; aangenomen dat het op dia 2 begint")

    # ratio uit direct opeenvolgende verzen (n -> n+1)
    afstanden = {b[0] - a[0] for a, b in zip(starts, starts[1:]) if b[1] - a[1] == 1}
    ratio = afstanden.pop() if len(afstanden) == 1 else None
    if len(afstanden) > 1:
        redenen.append("verzen hebben niet even veel dia's")  # onregelmatig gezang: geen aanvulling

    if ratio:
        # gaten opvullen waar een detectie gemist is (alleen als de afstand exact klopt)
        gevuld = [starts[0]]
        for a, b in zip(starts, starts[1:]):
            dn = b[1] - a[1]
            if dn > 1 and (b[0] - a[0]) == ratio * dn:
                for k in range(1, dn):
                    gevuld.append((a[0] + ratio * k, a[1] + k))
            elif dn > 1:
                redenen.append(f"versnummer springt van {a[1]} naar {b[1]}")
            gevuld.append(b)
        # doorlopen na de laatste detectie
        dia, nr = gevuld[-1]
        while dia + ratio <= laatste_dia:
            dia += ratio
            nr += 1
            gevuld.append((dia, nr))
        starts = gevuld
    else:
        for a, b in zip(starts, starts[1:]):
            if b[1] - a[1] > 1:
                redenen.append(f"versnummer springt van {a[1]} naar {b[1]}")
        if len(starts) == 1 and n_content_slides > 3:
            redenen.append("maar één vers herkend op een lang bestand")

    vertrouwen = "hoog" if not redenen else "laag"
    if not redenen and len(starts) == 1:
        vertrouwen = "gemiddeld"
    return _maak_resultaat(starts, laatste_dia, vertrouwen, redenen)


def _maak_resultaat(starts, laatste_dia, vertrouwen, redenen):
    verses = []
    for i, (dia, nr) in enumerate(starts):
        einde = (starts[i + 1][0] - 1) if i + 1 < len(starts) else laatste_dia
        if einde >= dia:
            verses.append({"vers": nr, "slides": list(range(dia, einde + 1))})
    lengtes = {len(v["slides"]) for v in verses}
    return {
        "slides_per_verse": lengtes.pop() if len(lengtes) == 1 else None,
        "n_verses": len(verses),
        "verses": verses,
        "conflicts": [],
        "confidence": vertrouwen,
        "opmerkingen": redenen,
        "bron": "ocr",
    }


def index_gezangen(bibliotheek_dir, output_dir, correcties=None, progress_cb=None):
    files = sorted(glob.glob(os.path.join(bibliotheek_dir, "**", "*.pptx"), recursive=True))
    files = [f for f in files if gezang_nummer(os.path.basename(f)) is not None]
    work_root = os.path.join(output_dir, "_render_gezangen")
    cache_root = os.path.join(output_dir, "render_cache")
    correcties = correcties or {}

    index, review = {}, []
    for i, pptx_path in enumerate(files, 1):
        base = os.path.splitext(os.path.basename(pptx_path))[0]
        nr = gezang_nummer(base)
        nr3 = f"{nr:03d}"
        key = f"{nr3}::{base}.pptx"
        safe = cache_mapnaam(base)
        try:
            dias = convert_pptx_to_slides(pptx_path, os.path.join(work_root, safe))
            detecties = []
            for pad in dias:
                idx = int(re.search(r"(\d+)", os.path.basename(pad)).group(1))
                if idx == 1:
                    continue
                num = detecteer_start(ocr_lines(pad))
                if num is not None:
                    detecties.append((idx, num))
            n_content = len(dias) - 1
            resultaat = bouw_gezang_index(n_content, detecties)
            resultaat["ocr_detecties"] = detecties
        except Exception as e:  # noqa: BLE001 -- één kapot bestand mag de rest niet stoppen
            index[key] = {"gezang": nr3, "file": base, "error": str(e)}
            review.append({"key": key, "reason": f"fout tijdens verwerken: {e}"})
            continue

        if nr3 in correcties or str(nr) in correcties:
            c = correcties.get(nr3) or correcties.get(str(nr))
            resultaat["verses"] = [{"vers": int(v), "slides": s} for v, s in sorted(c["verzen"].items(), key=lambda kv: int(kv[0]))]
            resultaat["n_verses"] = len(resultaat["verses"])
            lengtes = {len(v["slides"]) for v in resultaat["verses"]}
            resultaat["slides_per_verse"] = lengtes.pop() if len(lengtes) == 1 else None
            resultaat["confidence"] = "hoog"
            resultaat["bron"] = "handmatig gecontroleerd"
            resultaat["opmerkingen"] = [c["opmerking"]] if c.get("opmerking") else []

        resultaat.update({"gezang": nr3, "file": base, "n_content_slides": len(dias) - 1})
        index[key] = resultaat
        if resultaat["confidence"] == "laag":
            review.append({"key": key, "reason": "; ".join(resultaat["opmerkingen"]) or "laag vertrouwen"})

        # dia's naar de render_cache kopiëren (zelfde plek/vorm als de psalmen)
        doel = os.path.join(cache_root, safe)
        os.makedirs(doel, exist_ok=True)
        for pad in dias:
            shutil.copy2(pad, os.path.join(doel, os.path.basename(pad)))
        if progress_cb:
            progress_cb(key, i, len(files))

    return index, review


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("gebruik: python gezang_indexer.py <gezangen_map> <output_map (bv. data/)>")
        sys.exit(1)
    bib, uit = sys.argv[1], sys.argv[2]
    os.makedirs(uit, exist_ok=True)
    corr_pad = os.path.join(uit, "gezang_correcties.json")
    corr = json.load(open(corr_pad, encoding="utf-8")) if os.path.exists(corr_pad) else {}
    corr = {k: v for k, v in corr.items() if not k.startswith("_")}

    idx, rev = index_gezangen(bib, uit, correcties=corr,
                              progress_cb=lambda k, i, n: print(f"[{i}/{n}] {k}"))
    json.dump(idx, open(os.path.join(uit, "gezang_index.json"), "w", encoding="utf-8"),
              indent=1, ensure_ascii=False)
    json.dump(rev, open(os.path.join(uit, "review_needed_gezang.json"), "w", encoding="utf-8"),
              indent=1, ensure_ascii=False)
    shutil.rmtree(os.path.join(uit, "_render_gezangen"), ignore_errors=True)
    hoog = sum(1 for v in idx.values() if v.get("confidence") in ("hoog", "gemiddeld"))
    print(f"\nKlaar. {len(idx)} gezangen verwerkt, {hoog} betrouwbaar, {len(rev)} hebben een check nodig.")
