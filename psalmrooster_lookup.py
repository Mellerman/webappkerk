"""
psalmrooster_lookup.py
========================
Zoekt de voorzang op uit het jaarrooster (data/psalmrooster_2026.json),
op basis van datum + tijd + kerk. Dit lost de beperking op dat de
"voorzang" nooit in de liturgie zelf staat (die volgt namelijk uit een
vast jaarrooster, niet uit een losse keuze per dienst).

GEBRUIK:
    from psalmrooster_lookup import zoek_voorzang
    voorzang = zoek_voorzang("2026-08-09", "18:00", "MK")
    # -> "Ps 1: 2"  (of None als niet gevonden)
"""
import json, os

def _laad_rooster(data_dir):
    pad = os.path.join(data_dir, "psalmrooster_2026.json")
    if not os.path.exists(pad):
        return []
    return json.load(open(pad, encoding="utf-8"))


def zoek_voorzang(datum, tijd, kerk, data_dir=None):
    """datum: 'YYYY-MM-DD', tijd: 'HH:MM', kerk: 'GK' of 'MK'.
    Geeft de ruwe voorzang-referentie terug (bv. 'Ps 1: 2'), of None als
    er geen exacte match is (bv. datum/tijd/kerk komt niet voor in het
    rooster, of het is een bijzondere dienst zonder psalmverwijzing zoals
    kerstliederen)."""
    data_dir = data_dir or os.path.join(os.path.dirname(__file__), "data")
    rooster = _laad_rooster(data_dir)
    for record in rooster:
        if record["datum"] == datum and record["tijd"] == tijd and record["kerk"] == kerk:
            return record["voorzang"]
    return None
