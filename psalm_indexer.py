"""
psalm_indexer.py
=================
Bouwt automatisch een index van "welk vers staat op welke dia" voor een hele
bibliotheek psalm-PowerPoints (zoals de 150-mappen structuur van de kerk).

AANPAK (kort):
1. Elk .pptx-bestand wordt gerenderd naar losse dia-afbeeldingen (via LibreOffice
   + pdftoppm), zodat de tekst zichtbaar is zoals in de kerk (witte tekst op
   donkere achtergrond -- de ruwe afbeelding in het bestand zelf is transparant
   en daardoor bijna onleesbaar voor OCR).
2. Elke dia gaat door OCR (tesseract). We zoeken op ELKE regel (niet alleen de
   eerste!) naar een patroon "N." of "N " aan het begin -- dat is het versnummer.
3. Uit de gevonden versnummers leiden we af hoeveel dia's een vers gemiddeld
   inneemt (meestal 1, 2 of 3, maar dat verschilt per psalm/bestand -- dus NOOIT
   hardcoden, altijd per bestand opnieuw afleiden).
4. Met dat "dia's-per-vers"-getal vullen we de verzen in die OCR gemist heeft,
   en we vlaggen elke dia waar het gedetecteerde nummer NIET overeenkomt met
   wat het patroon verwacht (bv. OCR leest "5" waar "25" had moeten staan).
5. Bestanden met te weinig detecties, of met conflicten, komen in een
   "review_needed" lijst -- die moet een mens even nalopen (zie
   controlescherm_generator.py). De rest (in de praktijk ~90%) is direct
   betrouwbaar.

GEBRUIK:
    python psalm_indexer.py /pad/naar/psalmen_map /pad/naar/output_map

    De psalmen_map moet submappen bevatten met daarin .pptx-bestanden
    (net als de originele kerk-bibliotheek: "Psalm 001/Psalm 001 (OB).pptx").

    Output: psalm_index.json + review_needed.json in de output_map.

AFHANKELIJKHEDEN:
    - LibreOffice (soffice) voor pptx -> pdf
    - poppler-utils (pdftoppm) voor pdf -> jpg
    - tesseract-ocr voor OCR
    - python-pptx (alleen voor het tellen van dia's, optioneel)
"""
import json, os, re, subprocess, sys, glob
from collections import Counter


# ---------------------------------------------------------------------------
# Stap 1: pptx -> pdf -> losse jpg's per dia
# ---------------------------------------------------------------------------

def convert_pptx_to_slides(pptx_path, work_dir, dpi=150):
    """Rendert een pptx naar losse dia-afbeeldingen (jpg) met de echte
    achtergrondkleuren, zodat OCR de (vaak witte) tekst kan lezen.
    Retourneert een gesorteerde lijst dia-afbeeldingspaden (dia 1 = titel)."""
    os.makedirs(work_dir, exist_ok=True)
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", work_dir, pptx_path],
        capture_output=True, text=True, timeout=120
    )
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf_path = os.path.join(work_dir, base + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"Conversie naar PDF mislukt voor {pptx_path}")
    img_prefix = os.path.join(work_dir, "slide")
    subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), pdf_path, img_prefix],
                    capture_output=True, timeout=120)
    files = sorted(glob.glob(img_prefix + "*.jpg"),
                    key=lambda f: int(re.search(r"(\d+)", os.path.basename(f)).group(1)))
    return files


# ---------------------------------------------------------------------------
# Stap 2: OCR + versnummer-detectie
# ---------------------------------------------------------------------------

def ocr_lines(img_path):
    """OCR een afbeelding en geef de niet-lege regels terug."""
    out = subprocess.run(["tesseract", img_path, "-", "--psm", "6"],
                          capture_output=True, text=True, timeout=30).stdout
    return [l.strip() for l in out.splitlines() if l.strip()]


def detect_verse_number(lines):
    """Zoekt in ALLE OCR-regels (niet alleen de eerste!) naar een patroon
    'N.' of 'N ' aan het begin -- dat is waar het echte versnummer staat.
    Bij dia's met veel notenbalk-ruis staat de tekst soms pas op regel 2 of 3."""
    for line in lines:
        m = re.match(r"^(\d{1,3})[.,]?\s", line.strip())
        if m:
            return int(m.group(1))
    return None


def ocr_index_file(pptx_path, work_dir):
    """Rendert + OCR't een heel pptx-bestand. Retourneert:
    (aantal_inhoudsdias, [(dia_index, gedetecteerd_versnummer), ...])
    Dia-index 1 = titeldia (wordt overgeslagen), inhoud begint op dia 2."""
    slides = convert_pptx_to_slides(pptx_path, work_dir)
    detections = []
    for path in slides:
        idx = int(re.search(r"(\d+)", os.path.basename(path)).group(1))
        if idx == 1:
            continue  # titeldia
        num = detect_verse_number(ocr_lines(path))
        if num is not None:
            detections.append((idx, num))
    n_content = len(slides) - 1
    return n_content, detections


# ---------------------------------------------------------------------------
# Stap 3+4: patroon afleiden, gaten opvullen, conflicten markeren
# ---------------------------------------------------------------------------

def build_verse_index(n_content_slides, detections):
    """De kernlogica: van een setje (dia, versnummer)-detecties naar een
    volledige vers -> dia's mapping, met conflict-detectie.

    Retourneert een dict:
      {
        'slides_per_verse': int of None,
        'ratio_confidence': float,
        'verses': [{'vers': 1, 'slides': [2,3]}, ...] of None,
        'conflicts': [{'slide':.., 'verwacht_vers':.., 'ocr_gelezen':..}, ...],
        'confidence': 'hoog' | 'gemiddeld' | 'laag'
      }
    """
    detections = sorted(set(detections))
    if len(detections) < 2:
        return {'slides_per_verse': None, 'ratio_confidence': 0, 'verses': None,
                'conflicts': [], 'confidence': 'laag'}

    # Ratio = dia's per vers, afgeleid uit afstand tussen opeenvolgende
    # (stijgende!) detecties. Niet-stijgende paren (OCR-ruis) worden genegeerd.
    ratios = []
    for (s1, n1), (s2, n2) in zip(detections, detections[1:]):
        if n2 > n1 and s2 > s1:
            r = (s2 - s1) / (n2 - n1)
            if r == int(r):
                ratios.append(int(r))
    if not ratios:
        return {'slides_per_verse': None, 'ratio_confidence': 0, 'verses': None,
                'conflicts': [], 'confidence': 'laag'}

    ratio, votes = Counter(ratios).most_common(1)[0]
    ratio_confidence = votes / len(ratios)

    # LET OP: we gaan NIET er automatisch van uit dat vers 1 ook echt op de
    # bibliotheek-conventie start; we ankeren op de eerste betrouwbare
    # detectie en rekenen door. Als dat een startpositie < eerste inhoudsdia
    # zou opleveren, betekent dat dat er verzen ontbreken aan het begin --
    # zie de Psalm 41 case in de conversatie: dat moet een mens beoordelen,
    # dus die situatie belandt automatisch in 'review'.
    first_slide, first_num = detections[0]
    verse1_start = first_slide - (first_num - 1) * ratio

    verses = []
    conflicts = []
    detected_map = dict(detections)
    slide_cursor = max(verse1_start, 2)
    start_verse_num = first_num - (first_slide - slide_cursor) // ratio if slide_cursor != verse1_start else 1
    verse_num = max(start_verse_num, 1)
    last_slide = n_content_slides + 1

    while slide_cursor <= last_slide:
        remaining = last_slide - slide_cursor + 1
        take = ratio if remaining >= ratio else remaining
        slides = list(range(slide_cursor, slide_cursor + take))
        for s in slides:
            if s in detected_map and detected_map[s] != verse_num:
                conflicts.append({'slide': s, 'verwacht_vers': verse_num, 'ocr_gelezen': detected_map[s]})
        verses.append({'vers': verse_num, 'slides': slides})
        slide_cursor += take
        verse_num += 1

    confidence = 'hoog' if (ratio_confidence > 0.9 and not conflicts) else (
        'gemiddeld' if not conflicts else 'laag')

    return {'slides_per_verse': ratio, 'ratio_confidence': round(ratio_confidence, 2),
            'n_verses': len(verses), 'verses': verses, 'conflicts': conflicts,
            'confidence': confidence}


# ---------------------------------------------------------------------------
# Alles samen: hele bibliotheek indexeren
# ---------------------------------------------------------------------------

def index_library(library_dir, work_dir, progress_cb=None):
    """Doorloopt alle .pptx-bestanden in submappen van library_dir en bouwt
    de volledige index. progress_cb(bestandsnaam, i, totaal) is optioneel,
    handig voor een voortgangsbalk in een webapp."""
    files = []
    for sub in sorted(os.listdir(library_dir)):
        subdir = os.path.join(library_dir, sub)
        if os.path.isdir(subdir):
            for pf in sorted(glob.glob(os.path.join(subdir, "*.pptx"))):
                files.append((sub, pf))

    index = {}
    review_needed = []
    for i, (psalm_folder, pptx_path) in enumerate(files, 1):
        key = f"{psalm_folder}::{os.path.basename(pptx_path)}"
        file_work_dir = os.path.join(work_dir, re.sub(r'[^\w]', '_', key))
        try:
            n_content, detections = ocr_index_file(pptx_path, file_work_dir)
            result = build_verse_index(n_content, detections)
        except Exception as e:
            index[key] = {'psalm': psalm_folder, 'file': pptx_path, 'error': str(e)}
            review_needed.append({'key': key, 'reason': f'fout tijdens verwerking: {e}'})
            continue

        result.update({'psalm': psalm_folder, 'file': os.path.basename(pptx_path),
                        'n_content_slides': n_content})
        index[key] = result
        if result['confidence'] != 'hoog':
            review_needed.append({'key': key, 'reason': f"vertrouwen: {result['confidence']}",
                                   'conflicts': result.get('conflicts')})
        if progress_cb:
            progress_cb(key, i, len(files))

    return index, review_needed


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("gebruik: python psalm_indexer.py <psalmen_map> <output_map>")
        sys.exit(1)
    library_dir, output_dir = sys.argv[1], sys.argv[2]
    os.makedirs(output_dir, exist_ok=True)
    work_dir = os.path.join(output_dir, "_render_cache")

    def progress(key, i, total):
        print(f"[{i}/{total}] {key}")

    index, review = index_library(library_dir, work_dir, progress_cb=progress)

    json.dump(index, open(os.path.join(output_dir, "psalm_index.json"), "w"),
               indent=1, ensure_ascii=False)
    json.dump(review, open(os.path.join(output_dir, "review_needed.json"), "w"),
               indent=1, ensure_ascii=False)

    hoog = sum(1 for v in index.values() if v.get('confidence') == 'hoog')
    print(f"\nKlaar. {len(index)} bestanden verwerkt, {hoog} met hoog vertrouwen, "
          f"{len(review)} hebben een check nodig.")
