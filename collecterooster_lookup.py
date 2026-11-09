"""
collecterooster_lookup.py
===========================
Zoekt de 3 collectedoelen (rondgang 1, 2, 3) op voor een datum, gefilterd
op wijk 3-4. Rondgang 2 en 3 zijn altijd voor alle wijken (1-4) samen;
alleen rondgang 1 verschilt soms per wijk-groep (1-2 vs 3-4).

GEBRUIK:
    from collecterooster_lookup import zoek_collectedoelen
    doelen = zoek_collectedoelen("2026-08-09")
    # -> ["Evangelisatie", "Kerkrentmeesterlijk Beheer", "Huis van de Kerk"]
"""
import json, os

WIJK = "3-4"  # vast ingesteld: Hervormde Gemeente Sliedrecht wijk 3-4


def _laad_rooster(data_dir):
    pad = os.path.join(data_dir, "collecterooster_2026.json")
    if not os.path.exists(pad):
        return {}
    return json.load(open(pad, encoding="utf-8"))


def zoek_collectedoelen(datum, data_dir=None):
    """datum: 'YYYY-MM-DD'. Geeft een lijst van maximaal 3 collectedoelen
    terug (rondgang 1, 2, 3 op volgorde), of een lege lijst als de datum
    niet in het rooster staat."""
    data_dir = data_dir or os.path.join(os.path.dirname(__file__), "data")
    rooster = _laad_rooster(data_dir)
    regels = rooster.get(datum, [])
    if not regels:
        return []

    resultaat = []
    for rondgang_nr in (1, 2, 3):
        kandidaten = [r for r in regels if r["rondgang"] == rondgang_nr]
        if not kandidaten:
            continue
        # voorkeur: een regel die specifiek voor wijk 3-4 (of alle wijken
        # 1-4) geldt; zo niet, gewoon de eerste die er is (bv. bij een
        # datum waar rondgang 1 niet per wijk gesplitst is).
        match = next((r for r in kandidaten if r["wijk"] in (WIJK, "1-4")), kandidaten[0])
        resultaat.append(match["doel"])
    return resultaat
