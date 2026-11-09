"""
weerklank.py
============
Zet de zipjes van liedbundels online (Weerklank) om naar dia-afbeeldingen
voor de dienst-PowerPoint. Dit is de logica uit de losse app "kerkdia"
(zie kerkdia/app/app.py in de aangeleverde 7z), overgezet naar de vorm die
de hoofdapp nodig heeft: per Weerklank-lied één zipje, dat hier wordt
uitgepakt en per liedje een lijst klaargemaakte (bijgesneden) png-dia's
oplevert, klaar om achter de titel-dia van dat lied te plakken.

Hoe het werkt (getrouw aan kerkdia):
- De liedbundel-zipjes bevatten per vers 1 of meer .png's (namen als
  'projectie-388-muziek-1-1.png', 'projectie-WKL1-muziek-2-1.png',
  'nl-wk-projectiediap-p10-muziek-1-1.png' ...).
- Elke png is een witte/transparante achtergrond met de liedtekst erop.
  De witte rand wordt eraf geknipt (40 px marge, zelfde truc als kerkdia:
  helderheid wordt als doorzichtigheid behandeld, gecorrigeerd met de
  pixel linksboven) zodat de tekst groot op de dia kan.
- De png's worden in dienst-volgorde gesorteerd (zelfde sorteersleutel
  als kerkdia) en krijgen ieder een plaatsing op de dia mee: even groot
  als het scherm óf (voor smalle, hoge plaatjes) zo hoog als de dia, en
  opeenvolgende verzen houden dezelfde breedte zodat de tekst mooi
  doorloopt van dia naar dia.

GEBRUIK:
    from weerklank import verwerk_zip, plaatsing

    dia_plaatjes = verwerk_zip("lied_192.zip", werkmap)  # [{'pad','breedte','hoogte'}, ...]
    for info in dia_plaatjes:
        links, boven, breedte, hoogte, vorige = plaatsing(
            DIA_BREEDTE, DIA_HOOGTE, info, vorige)
"""
import os
import re
import zipfile

from PIL import Image, ImageChops
from werkzeug.utils import secure_filename

# Dia-afmetingen van het kerksjabloon (13,33 x 7,5 inch in EMU) -- dezelfde
# maten als add_picture_slide_after in pptx_surgery.py gebruikt.
DIA_BREEDTE = 12192000
DIA_HOOGTE = 6858000

_MARGE = 40  # px witte rand rond de tekst die meekomt (zoals kerkdia)


# ---------------------------------------------------------------------------
# Sorteren: zelfde sleutel als kerkdia's extract_song (is_sorting-variant),
# zodat de png's in dezelfde volgorde komen als de dia's die de losse
# kerkdia-app vroeger maakte.
# ---------------------------------------------------------------------------

def _sorteer_sleutel(bestandsnaam):
    naam = os.path.splitext(bestandsnaam)[0].lower()
    naam = re.sub(r"[_ ]+", "-", naam)   # 'projectie 192 muziek 1 1' / secure_filename-varianten
    naam = naam[3:]  # kerkdia haalt de eerste 3 tekens eraf ('pro' van 'projectie...')
    naam = naam.replace("projectiediap-", "")
    naam = naam.replace("muziek-", "")
    naam = naam.replace("wk-p", "wk-")
    delen = naam.split("-")
    sleutel = []
    songnummer = ""
    for num, deel in enumerate(delen):
        if num == 1:
            songnummer = deel
            deel = deel.strip("abcd")
        if deel.isdigit():
            sleutel.append(int(deel))
    while len(sleutel) < 3:
        sleutel.append(0)
    sleutel.append(songnummer)  # laatste (als tekst) voor stabiele sortering
    return tuple(sleutel)


# ---------------------------------------------------------------------------
# Bijsnijden (port van kerkdia's convert_image, zonder de dia-toevoeging)
# ---------------------------------------------------------------------------

def _bijsnijden(afbeeldingspad):
    """Knipt de witte/lege rand van een liedbundel-png af en geeft de
    bijgesneden PIL-afbeelding terug (of None als de png leeg is)."""
    afbeelding = Image.open(afbeeldingspad).convert("LA")
    alpha = afbeelding.getchannel("L")
    if alpha.getpixel((0, 0)) > 127:
        # witte achtergrond -> wit als doorzichtig behandelen
        alpha = ImageChops.invert(alpha)
    afbeelding.putalpha(alpha)
    kader = alpha.getbbox()
    if kader is None:
        return None
    links, boven, rechts, onder = kader
    links = max(links - _MARGE, 0)
    boven = max(boven - _MARGE, 0)
    rechts = min(rechts + _MARGE, afbeelding.width)
    onder = min(onder + _MARGE, afbeelding.height)
    afbeelding = afbeelding.crop((links, boven, rechts, onder))
    return afbeelding


_MAX_BESTANDEN = 400            # zo veel png's zit er nooit in een liedbundel-zipje
_MAX_TOTAAL_BYTES = 300 * 1024 * 1024   # bescherming tegen een 'zip-bom'


def verwerk_zip(zip_pad, werkmap):
    """Pakt een liedbundel-zipje uit en geeft de bijgesneden png-dia's terug
    in dienst-volgorde: een lijst dicts {'pad', 'breedte', 'hoogte'}.
    De png's worden naast de zip in `werkmap` weggeschreven.

    Werkt ook als de png's in een submap zitten (zip van een hele map), met
    spaties of hoofdletters in de naam; macOS-rommel (__MACOSX, ._bestand) en
    andere bestanden worden genegeerd. Er wordt nooit met de mapnaam uit de
    zip geschreven (geen zip-slip): alle png's komen plat in een eigen map."""
    os.makedirs(werkmap, exist_ok=True)
    uit_dir = os.path.join(werkmap, "extract")
    os.makedirs(uit_dir, exist_ok=True)
    png_namen = []
    with zipfile.ZipFile(zip_pad) as z:
        leden = [i for i in z.infolist()
                 if not i.is_dir()
                 and i.filename.lower().endswith(".png")
                 and "__macosx" not in i.filename.lower()
                 and not os.path.basename(i.filename).startswith(".")]
        if len(leden) > _MAX_BESTANDEN or sum(i.file_size for i in leden) > _MAX_TOTAAL_BYTES:
            raise ValueError("Het zipje bevat onverwacht veel of te grote bestanden")
        gebruikt = set()
        for i, lid in enumerate(leden):
            basis = os.path.basename(lid.filename.replace("\\", "/"))
            veilig = secure_filename(basis) or f"dia_{i}.png"
            if veilig in gebruikt:                       # zelfde naam in 2 submappen
                veilig = f"{i}_{veilig}"
            gebruikt.add(veilig)
            with z.open(lid) as bron, open(os.path.join(uit_dir, veilig), "wb") as doel:
                doel.write(bron.read())
            png_namen.append(veilig)
    if not png_namen:
        raise ValueError("Geen .png-bestanden gevonden in het zipje")

    resultaat = []
    for i, naam in enumerate(sorted(png_namen, key=_sorteer_sleutel), 1):
        afbeelding = _bijsnijden(os.path.join(uit_dir, naam))
        if afbeelding is None:
            continue  # volledig lege dia overslaan
        pad = os.path.join(werkmap, f"lied_dia_{i:03d}.png")
        afbeelding.save(pad, "PNG")
        resultaat.append({
            "pad": pad,
            "breedte": afbeelding.width,
            "hoogte": afbeelding.height,
        })
    if not resultaat:
        raise ValueError("Geen bruikbare lieddia's gevonden in het zipje")
    return resultaat


# ---------------------------------------------------------------------------
# Plaatsing op de dia (port van kerkdia's convert_image-plaatsingslogica)
# ---------------------------------------------------------------------------

def plaatsing(slide_breedte, slide_hoogte, info, vorige_breedte=0):
    """Berekent waar één lieddia-afbeelding op de dia komt: een tuple
    (links, boven, breedte, hoogte) in EMU. Opeenvolgende verzen van
    hetzelfde lied krijgen dezelfde breedte (zie kerkdia), zodat de tekst
    doorloopt. Retourneert ook de nieuwe 'vorige_breedte'."""
    beeld_verhouding = info["breedte"] / info["hoogte"]
    dia_verhouding = slide_breedte / slide_hoogte
    if vorige_breedte == 0 and dia_verhouding > beeld_verhouding:
        # hoog/smal plaatje: hele dia hoog, breedte volgt de verhouding
        breedte = round(slide_hoogte * beeld_verhouding)
        hoogte = slide_hoogte
    else:
        breedte = vorige_breedte if vorige_breedte else slide_breedte
        hoogte = round(breedte * info["hoogte"] / info["breedte"])
        if hoogte > slide_hoogte:
            breedte = round(breedte * slide_hoogte / hoogte)
            hoogte = slide_hoogte
    links = (slide_breedte - breedte) // 2
    boven = (slide_hoogte - hoogte) // 2
    return links, boven, breedte, hoogte, breedte
