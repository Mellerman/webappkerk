"""
collecterooster_parser.py
===========================
Zet het collecterooster (datum/rondgang/wijk -> collectedoel) om naar een
JSON die per datum de 3 rondgangen voor wijk 3-4 teruggeeft.

GEBRUIK:
    python collecterooster_parser.py collecterooster_ruw.txt collecterooster_2026.json
"""
import json, re, sys

REGEL_PATROON = re.compile(
    r'^(maandag|dinsdag|woensdag|donderdag|vrijdag|zaterdag|zondag)\s+'
    r'(\d{1,2})-(\d{1,2})-(\d{4})\s+'
    r'\d+\s+'                      # weeknummer
    r'(\d)\s+'                     # rondgang
    r'([\d\-]+)\s+'                # wijk (bv. '1-2', '3-4', '1-4')
    r'([\w\-]+)\s+'                # type collecte
    r'(.+)$',                      # collectedoel
    re.IGNORECASE)


def parse_collecterooster(ruwe_tekst):
    per_datum = {}
    for regel in ruwe_tekst.splitlines():
        regel = regel.strip()
        if not regel:
            continue
        m = REGEL_PATROON.match(regel)
        if not m:
            continue
        _, dag, maand, jaar, rondgang, wijk, type_collecte, doel = m.groups()
        datum_iso = f"{jaar}-{int(maand):02d}-{int(dag):02d}"
        per_datum.setdefault(datum_iso, []).append({
            'rondgang': int(rondgang), 'wijk': wijk,
            'type': type_collecte, 'doel': doel.strip(),
        })
    return per_datum


if __name__ == "__main__":
    ruwe_tekst = open(sys.argv[1], encoding='utf-8').read()
    per_datum = parse_collecterooster(ruwe_tekst)
    json.dump(per_datum, open(sys.argv[2], 'w', encoding='utf-8'), indent=1, ensure_ascii=False)
    print(f"{len(per_datum)} datums verwerkt -> {sys.argv[2]}")
