"""
psalmrooster_parser.py
========================
Zet het jaarrooster (datum/tijd/kerk/wijk -> voorzang) om naar een simpele
JSON-lijst die de webapp kan opzoeken.

GEBRUIK:
    python psalmrooster_parser.py rooster_ruw.txt psalmrooster_2026.json
"""
import json, re, sys

MAANDEN = {
    'januari': 1, 'februari': 2, 'maart': 3, 'april': 4, 'mei': 5, 'juni': 6,
    'juli': 7, 'augustus': 8, 'september': 9, 'oktober': 10, 'november': 11, 'december': 12,
}

DAG_PATROON = re.compile(
    r'(maandag|dinsdag|woensdag|donderdag|vrijdag|zaterdag|zondag)\s+'
    r'(\d{1,2})\s+(' + '|'.join(MAANDEN) + r')\s+(\d{4})',
    re.IGNORECASE)

TIJD_PATROON = re.compile(r'(\d{1,2})[:.](\d{2})')
KERK_PATROON = re.compile(r'\b(GK|MK)\s*(\d)?')
# de voorzang staat na de laatste 'K ' of 'B ' (Lit.-kolom) in het record
VOORZANG_PATROON = re.compile(
    r'\b[KB]\s+((?:W)?(?:Ps|Gz|Ld)\.?\s*\d{1,3}\s*:\s*\d{1,3}[a-z]?|[A-Za-z][\w\s]{3,30})\s*$')


def parse_rooster(ruwe_tekst):
    matches = list(DAG_PATROON.finditer(ruwe_tekst))
    records = []
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(ruwe_tekst)
        chunk = ruwe_tekst[start:end]

        dag_nr, maand_naam, jaar = m.group(2), m.group(3).lower(), m.group(4)
        maand = MAANDEN[maand_naam]
        datum_iso = f"{jaar}-{maand:02d}-{int(dag_nr):02d}"

        tijd_m = TIJD_PATROON.search(chunk)
        tijd = f"{tijd_m.group(1).zfill(2)}:{tijd_m.group(2)}" if tijd_m else None

        kerk_m = KERK_PATROON.search(chunk)
        kerk = kerk_m.group(1) if kerk_m else None
        wijk = kerk_m.group(2) if kerk_m and kerk_m.group(2) else None

        vz_m = VOORZANG_PATROON.search(chunk.strip())
        voorzang = vz_m.group(1).strip() if vz_m else None

        if tijd and kerk:
            records.append({
                'datum': datum_iso, 'tijd': tijd, 'kerk': kerk, 'wijk': wijk,
                'voorzang': voorzang,
            })
    return records


if __name__ == "__main__":
    ruwe_tekst = open(sys.argv[1], encoding='utf-8').read()
    records = parse_rooster(ruwe_tekst)
    json.dump(records, open(sys.argv[2], 'w', encoding='utf-8'), indent=1, ensure_ascii=False)
    print(f"{len(records)} diensten verwerkt -> {sys.argv[2]}")
    ontbrekend = [r for r in records if not r['voorzang']]
    print(f"{len(ontbrekend)} zonder herkende voorzang (bv. 'Kerstliederen', 'Laudate Deum' -- geen psalmverwijzing)")
