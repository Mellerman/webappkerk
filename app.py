"""
app.py
======
De webschil om alles heen: upload een liturgie, controleer/vul de velden
aan die niet uit het document te halen zijn, en download de PowerPoint.

Bewust STATELESS (geen database, geen sessies) -- fijn voor een simpele
homeserver-deployment. De geparste liturgie reist als verborgen JSON-veld
mee in het controle-formulier naar de volgende stap.

STARTEN:
    pip install -r requirements.txt
    python app.py
    -> http://localhost:5000

DATAMAP (data/) MOET BEVATTEN (zie README.md):
    psalm_index.json
    template_map.json
    sjabloon_avonddienst.pptx
    render_cache/<psalmmap>/slide-NN.jpg   (de voorgerenderde dia's)
    assets/orgel_grotekerk.png, assets/orgel_maranathakerk.png
"""
import os
import re
import json
import shutil
import tempfile
import time
import traceback
import uuid
from datetime import datetime
from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from assemble import assemble_dienst, _psalm_entry
from bijbeltekst import (STANDAARD_VERTALING, splits_verwijzingen, verwijzing_ophalen,
                         verwijzing_parsen, verwijzing_titeltekst)
from dienstverslag import schrijf_dienstverslag
from liturgie_parser import (herstel_kern_thema, parse_liturgie,
                             parse_muziekmomenten, parse_psalmregel)
from psalmrooster_lookup import zoek_voorzang, _laad_rooster as _laad_psalmrooster
from collecterooster_lookup import zoek_collectedoelen
from weerklank import verwerk_zip

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")     # altijd t.o.v. de app-map, nooit t.o.v. de werkmap
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Bewaar downloads zo veel dagen (elke dienst is ~5 MB); de verslagen blijven staan.
OUTPUT_BEWAARDAGEN = int(os.environ.get("OUTPUT_BEWAARDAGEN", "30"))


def _service_profiles():
    pad = os.path.join(DATA_DIR, "service_profiles.json")
    try:
        return json.load(open(pad, encoding="utf-8"))
    except (OSError, ValueError):
        return {"standard": {"label": "Standaarddienst", "kind": "standard"}}


def _geheime_sleutel():
    """FLASK_SECRET_KEY uit de omgeving, anders een willekeurige sleutel die één keer in
    .secret_key wordt bewaard (dus ook bij meerdere gunicorn-workers en na een herstart
    hetzelfde, en niet meer de openbare 'dev-secret')."""
    sleutel = os.environ.get("FLASK_SECRET_KEY")
    if sleutel:
        return sleutel
    pad = os.path.join(BASE_DIR, ".secret_key")
    try:
        if os.path.exists(pad):
            return open(pad, encoding="utf-8").read().strip()
        sleutel = uuid.uuid4().hex + uuid.uuid4().hex
        with open(pad, "w", encoding="utf-8") as f:
            f.write(sleutel)
        return sleutel
    except OSError:
        return uuid.uuid4().hex


app = Flask(__name__)
app.secret_key = _geheime_sleutel()
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024   # liturgie + max. 6 liedbundel-zipjes


def _ruim_output_op():
    """Verwijdert oude dienst-pptx'en uit output/ (de verslagen blijven)."""
    grens = time.time() - OUTPUT_BEWAARDAGEN * 86400
    try:
        for naam in os.listdir(OUTPUT_DIR):
            pad = os.path.join(OUTPUT_DIR, naam)
            if naam.lower().endswith(".pptx") and os.path.isfile(pad) and os.path.getmtime(pad) < grens:
                os.remove(pad)
    except OSError:
        pass

# De vaste volgorde van de dienst-onderdelen (zelfde labels als de webpagina).
SECTIE_VOLGORDE = [
    ("voorzang", "Voorzang"),
    ("opening", "Opening"),
    ("na_wet", "Na wet/apostolicum"),
    ("voor_prediking", "Voor de prediking"),
    ("na_prediking", "Na de prediking"),
    ("slot", "Slot"),
]


def regel_tekst(parsed):
    """Zet een geparste psalm-/liedregel terug om naar de tekst die je in het
    bewerkveld op het controlescherm ziet (en die parse_psalmregel weer
    dezelfde regel van maakt), bv. 'Psalm 19: 3, 4her' of 'Psalm 43: 3 (Wk)'."""
    verzen = ", ".join(str(v) + ("her" if v in parsed.get("herhaal", []) else "")
                       for v in parsed.get("verzen", []))
    nummer = str(parsed.get("nummer", "0"))
    try:
        nummer = str(int(nummer))
    except ValueError:
        pass
    nummer += str(parsed.get("suffix", ""))
    tekst = f"{parsed.get('soort', 'Lied')} {nummer}: {verzen}"
    if parsed.get("bundel") == "Wk" and parsed.get("soort") in ("Psalm", "Gezang"):
        tekst += " (Wk)"
    return tekst


def regel_label(parsed):
    """Weergave zoals op het controlescherm: 'Psalm 19 — 3, 4'."""
    verzen = ", ".join(str(v) for v in parsed.get("verzen", []))
    nummer = str(parsed.get("nummer", "0"))
    try:
        nummer = str(int(nummer))
    except ValueError:
        pass
    nummer += str(parsed.get("suffix", ""))
    return f"{parsed.get('soort', 'Lied')} {nummer}" + (f" — {verzen}" if verzen else "")


app.jinja_env.globals["regel_tekst"] = regel_tekst

_INDEX_CACHE = {}


def _laad_index(bestand):
    if bestand not in _INDEX_CACHE:
        pad = os.path.join(BASE_DIR, "data", bestand)
        _INDEX_CACHE[bestand] = json.load(open(pad, encoding="utf-8")) if os.path.exists(pad) else None
    return _INDEX_CACHE[bestand]


def _bibliotheek_status(parsed):
    """Voor een psalm/gezang uit de eigen bibliotheek: welke verzen zitten er
    niet in? None als het een Weerklank-lied is (die komen uit een zipje)."""
    if parsed.get("bundel") != "OB":
        return None
    index = _laad_index("gezang_index.json" if parsed["soort"] == "Gezang" else "psalm_index.json")
    if index is None:
        return {"ok": False, "ontbreekt": list(parsed["verzen"]), "reden": "index niet gevonden"}
    ontbreekt = []
    for vn in parsed["verzen"]:
        entry = _psalm_entry(index, parsed["nummer"], [vn])
        if entry is None or vn not in {v["vers"] for v in entry["verses"]}:
            ontbreekt.append(vn)
    return {"ok": not ontbreekt, "ontbreekt": ontbreekt}


@app.route("/parse_regel", methods=["POST"])
def parse_regel():
    """Controleert 1 handmatig ingetypte psalm-/liedregel voor het controlescherm."""
    tekst = ((request.get_json(silent=True) or {}).get("tekst") or "").strip()
    if not tekst:
        return jsonify(ok=True, leeg=True)
    if soort == "lezing" and re.search(r"(?:catechismus|\bcat\.?\s*zondag|heid\.?\s*cat\.?)", tekst, re.IGNORECASE):
        return jsonify(ok=True, leeg=False, resultaten=[{
            "tekst": tekst, "herkend": True, "titel": tekst,
            "deels": False, "automatisch": False, "referentie": True,
        }])
    parsed = parse_psalmregel(tekst)
    if not parsed:
        return jsonify(ok=False, leeg=False)
    return jsonify(ok=True, leeg=False, label=regel_label(parsed), bundel=parsed["bundel"],
                   soort=parsed["soort"], bibliotheek=_bibliotheek_status(parsed))


def normaliseer_tijd(tekst):
    """'9.30', '9:30 uur', '10u', '10' -> '09:30' / '10:00'; None als het geen tijd is."""
    m = re.match(r"^\s*(\d{1,2})(?:\s*[:.hu]\s*(\d{2}))?\s*(?:uur|u)?\s*$", tekst or "", re.IGNORECASE)
    if not m or int(m.group(1)) > 23 or int(m.group(2) or 0) > 59:
        return None
    return f"{int(m.group(1)):02d}:{m.group(2) or '00'}"


def _rooster_einddatums():
    """(laatste datum psalmrooster, laatste datum collecterooster) als 'YYYY-MM-DD'
    (of None): zo kan de app melden dat een jaarrooster op is."""
    psalm = [r["datum"] for r in _laad_psalmrooster(DATA_DIR)]
    pad = os.path.join(DATA_DIR, "collecterooster_2026.json")
    collecte = []
    if os.path.exists(pad):
        try:
            collecte = list(json.load(open(pad, encoding="utf-8")).keys())
        except (OSError, ValueError):
            pass
    return (max(psalm) if psalm else None, max(collecte) if collecte else None)


def _voorzang_uit_rooster(datum, aanvangstijd, kerk):
    """Zoekt de voorzang op in het jaarrooster. Geeft (geparste_regel | None, melding | None).
    De melding legt uit WAAROM er niets gevonden is (rooster op, geen psalm, niet in rooster)."""
    tijd = normaliseer_tijd(aanvangstijd)
    if not (datum and tijd):
        return None, None
    einde, _ = _rooster_einddatums()
    if einde and datum > einde:
        return None, (f"Het psalmrooster loopt tot {einde}: voor {datum} is geen voorzang opgezocht. "
                      "Laad het nieuwe jaarrooster in (psalmrooster_parser.py, zie README).")
    ruw = zoek_voorzang(datum, tijd, kerk, data_dir=DATA_DIR)
    if not ruw:
        return None, (f"Geen voorzang gevonden in het rooster voor {datum} om {tijd} "
                      f"({'Maranathakerk' if kerk == 'MK' else 'Grote Kerk'}): voorzang blijft leeg.")
    geparsed = parse_psalmregel(ruw)
    if not geparsed:
        return None, (f"De voorzang in het rooster ('{ruw}') is geen psalm of gezang dat automatisch "
                      "ingevuld kan worden (bv. kerstliederen): voeg hem zelf toe.")
    geparsed["uit_rooster"] = True
    return geparsed, None


@app.route("/parse_bijbel", methods=["POST"])
def parse_bijbel():
    """Controleert 1 bijbelverwijzing (schriftlezing of kerntekst) voor het
    controlescherm: wordt hij herkend, en kan de tekst automatisch ingevuld worden?"""
    data = request.get_json(silent=True) or {}
    tekst = (data.get("tekst") or "").strip()
    soort = data.get("soort") or "lezing"
    vertaling = data.get("vertaling") or STANDAARD_VERTALING
    if not tekst:
        return jsonify(ok=True, leeg=True)
    # meerdere verwijzingen in 1 veld ('Genesis 1: 1-5 en Marcus 4: 35-41') worden
    # bij het downloaden gesplitst in aparte schriftlezingen
    delen = splits_verwijzingen(tekst) if soort == "lezing" else [tekst]
    resultaten = []
    for deel in delen:
        v = verwijzing_parsen(deel)
        if not v:
            resultaten.append({"tekst": deel, "herkend": False})
            continue
        verzen = verwijzing_ophalen(vertaling, v, met_delen=(soort == "kerntekst"))
        aantal = sum(len(x) for x in verzen.values()) if verzen else 0
        resultaten.append({"tekst": deel, "herkend": True,
                           "titel": verwijzing_titeltekst(v) if v["heel"] else deel,
                           "deels": not v["heel"], "automatisch": bool(verzen), "aantal": aantal})
    return jsonify(ok=all(r["herkend"] for r in resultaten), leeg=False, resultaten=resultaten)


@app.route("/rooster_info")
def rooster_info():
    """Welke diensten staan er op deze datum in het psalmrooster (kerk, tijd,
    voorzang) en wat zijn de collectedoelen? Voor de knoppen onder het datumveld."""
    datum = (request.args.get("datum") or "").strip()
    if not datum:
        return jsonify(opties=[], collecte=[])
    opties = [{"kerk": r["kerk"], "tijd": r["tijd"], "voorzang": r.get("voorzang")}
              for r in _laad_psalmrooster(DATA_DIR) if r["datum"] == datum]
    opties.sort(key=lambda o: o["tijd"])
    return jsonify(opties=opties, collecte=zoek_collectedoelen(datum, data_dir=DATA_DIR))


def label_van(naam):
    """'voor_prediking' -> 'Voor de prediking'."""
    return dict(SECTIE_VOLGORDE).get(naam, naam)


def _weerklank_rijen(liturgie):
    """Welke Weerklank-liederen/psalmen (bundel 'Wk') staan in de liturgie?
    Geeft per plek: (sectie_naam, dienstdeel-label, lied-tekst)."""
    rijen = []
    for i, regel in enumerate(liturgie.get("muziekmomenten") or []):
        if not regel or regel.get("bundel") != "Wk":
            continue
        label = regel.get("label") or f"Liedmoment {i + 1}"
        rijen.append({"naam": f"moment_{i}", "dienstdeel": label,
                      "lied": regel_tekst(regel), "index": i})
    for naam, label in SECTIE_VOLGORDE:
        regel = (liturgie.get("secties") or {}).get(naam)
        if not regel or regel.get("bundel") != "Wk":
            continue
        verzen = [str(v) + ("her" if v in regel.get("herhaal", []) else "")
                  for v in regel.get("verzen", [])]
        nummer = regel_tekst(regel).split(":", 1)[0].split(" ", 1)[-1]
        lied = f"{regel.get('soort', 'Lied')} {nummer}: {', '.join(verzen)}"
        rijen.append({"naam": naam, "dienstdeel": label, "lied": lied})
    return rijen


@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/parse", methods=["POST"])
def parse():
    bestand = request.files.get("liturgie")
    if not bestand or not bestand.filename:
        flash("Kies eerst een liturgie-bestand (.docx).")
        return redirect(url_for("index"))

    stam, ext = os.path.splitext(bestand.filename)
    if ext.lower() != ".docx":
        flash(f"'{bestand.filename}' is geen .docx-bestand. Sla de liturgie in Word op als .docx "
              "(Bestand > Opslaan als) en probeer het opnieuw.")
        return redirect(url_for("index"))

    # Eigen tijdelijke map per upload: twee mensen tegelijk met dezelfde bestandsnaam
    # overschrijven (en wissen) elkaars bestand niet meer. De naam blijft bewaard omdat
    # de voorganger soms ook in de bestandsnaam staat. secure_filename maakt van een
    # naam met alleen bijzondere tekens niets meer, dan vallen we terug op 'liturgie'.
    upload_map = tempfile.mkdtemp(prefix="liturgie_", dir=UPLOAD_DIR)
    upload_path = os.path.join(upload_map, (secure_filename(stam) or "liturgie") + ".docx")
    try:
        bestand.save(upload_path)
        liturgie = parse_liturgie(upload_path)
    except Exception as e:
        app.logger.error("parse_liturgie failed for %s: %s", bestand.filename, traceback.format_exc())
        flash(f"Kon het document niet lezen: {e}")
        return redirect(url_for("index"))
    finally:
        shutil.rmtree(upload_map, ignore_errors=True)

    # Voorzang alvast uit het rooster halen, zodat hij op het controlescherm staat (en
    # je bij een Weerklank-voorzang meteen het zipje kunt uploaden). Wijzig je later
    # kerk of tijd, dan wordt hij bij het downloaden opnieuw opgezocht (zie /generate).
    if liturgie.get("datum") and liturgie.get("aanvangstijd") and liturgie.get("kerk") \
            and "voorzang" not in (liturgie.get("secties") or {}):
        voorzang, _ = _voorzang_uit_rooster(liturgie["datum"], liturgie["aanvangstijd"], liturgie["kerk"])
        if voorzang:
            liturgie.setdefault("secties", {})["voorzang"] = voorzang
            if liturgie.get("dienstprofiel") in ("avondmaal", "avondmaal_voortzetting"):
                liturgie.setdefault("muziekmomenten", []).insert(0, voorzang)

    return render_template(
        "review.html",
        liturgie=liturgie,
        bestandsnaam=bestand.filename,
        weerklank_rijen=_weerklank_rijen(liturgie),
        sectie_volgorde=SECTIE_VOLGORDE,
        wk_namen=[r["naam"] for r in _weerklank_rijen(liturgie)],
        muziek_weergave=liturgie.get("muziekmomenten") or [],
        dienstprofielen=_service_profiles(),
        gekozen_profiel=liturgie.get("dienstprofiel", "standard"),
    )


@app.route("/generate", methods=["POST"])
def generate():
    liturgie = json.loads(request.form["liturgie_json"])
    service_profile = request.form.get("service_profile", "standard")
    profiles = _service_profiles()
    if service_profile not in profiles:
        service_profile = "standard"
    profile_cfg = profiles.get(service_profile, {})
    liturgie["dienstprofiel"] = service_profile

    if profile_cfg.get("kind") == "voorbeeld":
        muziek = []
        for ruw in request.form.getlist("muziekmomenten[]"):
            parsed_lijst = parse_muziekmomenten(ruw.strip()) if ruw.strip() else []
            muziek.append(parsed_lijst[0] if parsed_lijst else None)
        liturgie["muziekmomenten"] = muziek

    kerk = request.form.get("kerk", "GK")
    type_dienst_override = request.form.get("type_dienst")
    if type_dienst_override:
        liturgie["type_dienst"] = type_dienst_override

    voorganger = request.form.get("voorganger", "").strip() or None
    organist = request.form.get("organist", "").strip() or None
    aanvangstijd = request.form.get("aanvangstijd", "").strip() or None
    aanvangstijd_ruw = aanvangstijd
    # Alleen een geldige tijd (HH:MM) gaat verder; een onleesbare tijd ('half tien')
    # wordt weggelaten (vroeger: '00:00' op de titeldia of een crash).
    aanvangstijd = normaliseer_tijd(aanvangstijd) if aanvangstijd else None
    datum = request.form.get("datum", "").strip() or None

    # Bevindingen tijdens het samenstellen worden na afloop naar een
    # dienstverslag op schijf geschreven (buiten de webapp om) en staan op
    # het scherm na het downloaden. Wat de parser al opmerkte (te veel/weinig
    # liederen, kerntekst en thema gewisseld, ...) staat hier ook bij.
    bevindingen = list(liturgie.get("parse_bevindingen") or [])

    if aanvangstijd_ruw and not aanvangstijd:
        bevindingen.append(f"Aanvangstijd '{aanvangstijd_ruw}' niet herkend als tijd (bv. 09:30): "
                           "weggelaten op de titeldia; voorzang uit het rooster kan daardoor niet "
                           "gezocht worden.")

    # Een voorzang die bij het inlezen uit het rooster is gehaald, zoeken we nu opnieuw op
    # met de definitieve kerk/datum/tijd (die kun je op het controlescherm nog gewijzigd
    # hebben). Handmatig aangepast ('Wijzig') blijft staan.
    if (liturgie.get("secties") or {}).get("voorzang", {}).get("uit_rooster") \
            and "regel_voorzang" not in request.form:
        liturgie["secties"].pop("voorzang")

    # Type dienst bepaalt of de Wet van God (ochtend) of de geloofsbelijdenis
    # (avond) getoond wordt. Onbekend -> uit de aanvangstijd afleiden, en zo
    # niet: expliciet melden in plaats van stil een keuze te maken.
    if not liturgie.get("type_dienst"):
        if aanvangstijd and normaliseer_tijd(aanvangstijd):
            liturgie["type_dienst"] = "ochtend" if int(aanvangstijd[:2]) < 13 else "avond"
            bevindingen.append(f"Type dienst niet ingevuld: uit de aanvangstijd afgeleid als "
                               f"{liturgie['type_dienst']}dienst.")
        else:
            bevindingen.append("Type dienst niet ingevuld en niet herkend: behandeld als avonddienst "
                               "(geloofsbelijdenis i.p.v. Wet van God). Controleer dit!")

    # Psalmen/liederen die op het controlescherm handmatig zijn aangepast
    # (veld 'regel_<onderdeel>'; alleen aanwezig als de gebruiker op "Wijzig"
    # heeft geklikt). Leeg veld = dat onderdeel weglaten. De regel wordt hier
    # opnieuw door dezelfde parser gehaald als bij het inlezen van de liturgie.
    liturgie.setdefault("secties", {})
    for naam, label in SECTIE_VOLGORDE:
        veld = f"regel_{naam}"
        if veld not in request.form:
            continue
        nieuw = request.form.get(veld, "").strip()
        oud = liturgie["secties"].get(naam)
        oud_tekst = regel_tekst(oud) if oud else "(geen)"
        if not nieuw:
            if oud:
                liturgie["secties"].pop(naam, None)
                bevindingen.append(f"Handmatig aangepast: {label}: '{oud_tekst}' weggelaten.")
            continue
        geparsed = parse_psalmregel(nieuw)
        if not geparsed:
            bevindingen.append(f"Handmatige regel voor {label} ('{nieuw}') niet herkend; "
                               f"oorspronkelijke waarde ('{oud_tekst}') gehouden.")
            continue
        liturgie["secties"][naam] = geparsed
        if regel_tekst(geparsed) != oud_tekst:
            bevindingen.append(f"Handmatig aangepast: {label}: '{oud_tekst}' -> '{regel_tekst(geparsed)}'.")

    # Schriftlezingen en kerntekst kunnen op het controlescherm aangepast
    # worden (de geparste versie staat in liturgie_json).
    if "lezingen[]" in request.form:
        ingevulde_lezingen = []
        for veld in request.form.getlist("lezingen[]"):
            for verw in splits_verwijzingen(veld.strip()):
                if verw not in ingevulde_lezingen:
                    ingevulde_lezingen.append(verw)
        liturgie["schriftlezingen"] = ingevulde_lezingen
    if "thema" in request.form:
        liturgie["thema"] = request.form.get("thema", "").strip() or None
    if "kerntekst" in request.form:
        kerntekst_veld = request.form.get("kerntekst", "").strip()
        liturgie["kerntekst"] = kerntekst_veld or None
    # thema en kerntekst verwisseld (bv. een bijbelverwijzing in het themaveld)?
    bevindingen.extend(herstel_kern_thema(liturgie))

    # Bijbelvertaling voor het automatisch invullen van de schriftlezingen
    # (en kerntekst). De gemeente gebruikt standaard de HSV.
    bijbel_vertaling = request.form.get("bijbel_vertaling", "").strip() or STANDAARD_VERTALING

    # Collectedoelen komen niet uit een invulveld, maar worden -- net als de
    # voorzang -- automatisch opgezocht op basis van de ingevulde datum, gefilterd
    # op wijk 3-4 (zie collecterooster_lookup.py). Het pad is t.o.v. de app-map:
    # vroeger stond hier het relatieve "data", dat alleen werkte als de app vanuit
    # de eigen map gestart was (niet onder systemd/gunicorn).
    collecte_opgezocht = bool(datum)
    collectedoelen = zoek_collectedoelen(datum, data_dir=DATA_DIR) if datum else None

    voorzang_opgezocht = bool(datum and aanvangstijd
                             and "voorzang" not in liturgie.get("secties", {}))

    # Voorzang staat nooit in de liturgie zelf (die volgt uit een vast
    # jaarrooster). Als de gebruiker een datum invult EN er nog geen
    # voorzang gevonden is, proberen we 'm op te zoeken in het rooster.
    if voorzang_opgezocht:
        voorzang, melding = _voorzang_uit_rooster(datum, aanvangstijd, kerk)
        if voorzang:
            liturgie.setdefault("secties", {})["voorzang"] = voorzang
        elif melding:
            app.logger.info(melding)
            bevindingen.append(melding)

    _psalm_einde, _collecte_einde = _rooster_einddatums()
    if not datum:
        bevindingen.append("Geen datum ingevuld: voorzang en collectedoelen zijn niet automatisch "
                           "opgezocht.")
    elif not collectedoelen:
        if _collecte_einde and datum > _collecte_einde:
            bevindingen.append(f"Het collecterooster loopt tot {_collecte_einde}: voor {datum} zijn geen "
                               "collectedoelen opgezocht. Laad het nieuwe jaarrooster in "
                               "(collecterooster_parser.py, zie README).")
        else:
            bevindingen.append(f"Geen collectedoelen gevonden in het rooster voor {datum}.")
    if datum and not aanvangstijd and not aanvangstijd_ruw and "voorzang" not in liturgie.get("secties", {}):
        bevindingen.append("Geen aanvangstijd ingevuld: voorzang is niet uit het rooster gezocht.")

    # unieke naam (ook als twee mensen in dezelfde seconde genereren) en oude downloads opruimen
    _ruim_output_op()
    bestandsnaam = f"dienst_{kerk}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}.pptx"
    output_pad = os.path.join(OUTPUT_DIR, bestandsnaam)

    # -- Weerklank-liedbundel-zipjes ----------------------------------------
    # Voor elk Weerklank-lied/psalm in de liturgie kan op het controlescherm
    # het liedbundel-zipje van liedbundels online zijn geüpload. Die zetten
    # we hier om naar bijgesneden png-dia's (port van de losse kerkdia-app,
    # zie weerklank.py) en geven ze mee aan de samenstelling. Zonder zipje
    # blijft de titel-placeholder-dia staan, zoals voorheen.
    weerklank_afbeeldingen = {}
    weerklank_fouten = {}       # sectie_naam -> waarom het zipje niet verwerkt kon worden
    werkmap = None
    if profile_cfg.get("kind") == "voorbeeld" and any(
            r and r.get("bundel") == "Wk" for r in liturgie.get("muziekmomenten", [])):
        werkmap = tempfile.mkdtemp(prefix="weerklank_")
        for i, regel in enumerate(liturgie.get("muziekmomenten", [])):
            if not regel or regel.get("bundel") != "Wk":
                continue
            zip_veld = request.files.get(f"wkzip_moment_{i}")
            if not zip_veld or not zip_veld.filename:
                continue
            if not zip_veld.filename.lower().endswith(".zip"):
                weerklank_fouten[f"moment_{i}"] = f"'{zip_veld.filename}' is geen .zip-bestand"
                bevindingen.append(f"Weerklank liedmoment {i + 1}: '{zip_veld.filename}' is geen .zip-bestand; "
                                   "alleen de titel-dia is neergezet.")
                continue
            zip_pad = os.path.join(werkmap, secure_filename(zip_veld.filename))
            zip_veld.save(zip_pad)
            try:
                plaatjes = verwerk_zip(zip_pad, os.path.join(werkmap, f"wk_moment_{i}"))
                if plaatjes:
                    weerklank_afbeeldingen[str(i)] = plaatjes
            except Exception as e:
                weerklank_fouten[f"moment_{i}"] = str(e)
                bevindingen.append(f"Weerklank liedmoment {i + 1}: het zipje kon niet verwerkt worden "
                                   f"({e}); alleen de titel-dia is neergezet.")
    elif profile_cfg.get("kind") != "voorbeeld" and any(r["naam"] for r in _weerklank_rijen(liturgie)):
        werkmap = tempfile.mkdtemp(prefix="weerklank_")
        for naam, _label in SECTIE_VOLGORDE:
            regel = (liturgie.get("secties") or {}).get(naam)
            if not regel or regel.get("bundel") != "Wk":
                continue
            zip_veld = request.files.get(f"wkzip_{naam}")
            if not zip_veld or not zip_veld.filename:
                continue  # geen zipje geüpload -> alleen de titel-dia
            if not zip_veld.filename.lower().endswith(".zip"):
                app.logger.warning("Weerklank '%s': '%s' is geen .zip-bestand; alleen titel-dia.",
                                   naam, zip_veld.filename)
                weerklank_fouten[naam] = f"'{zip_veld.filename}' is geen .zip-bestand"
                bevindingen.append(f"Weerklank {label_van(naam)}: '{zip_veld.filename}' is geen .zip-bestand; "
                                   "alleen de titel-dia is neergezet.")
                continue
            zip_pad = os.path.join(werkmap, secure_filename(zip_veld.filename))
            zip_veld.save(zip_pad)
            try:
                plaatjes = verwerk_zip(zip_pad, os.path.join(werkmap, f"wk_{naam}"))
            except Exception as e:
                app.logger.warning("Weerklank '%s': '%s' kon niet verwerkt worden (%s); alleen titel-dia.",
                                   naam, zip_veld.filename, e)
                weerklank_fouten[naam] = str(e)
                bevindingen.append(f"Weerklank {label_van(naam)}: het zipje '{zip_veld.filename}' kon niet "
                                   f"verwerkt worden ({e}); alleen de titel-dia is neergezet.")
                continue
            if plaatjes:
                weerklank_afbeeldingen[naam] = plaatjes

    pptx_werkmap = tempfile.mkdtemp(prefix="pptx_werk_")
    try:
        assemble_dienst(
            work_dir=pptx_werkmap,
            liturgie=liturgie,
            kerk=kerk,
            output_path=output_pad,
            organist=organist,
            voorganger_override=voorganger,
            aanvangstijd=aanvangstijd,
            collectedoelen=collectedoelen,
            bijbel_vertaling=bijbel_vertaling,
            weerklank_afbeeldingen=weerklank_afbeeldingen or None,
            bevindingen=bevindingen,
            service_profile=service_profile,
        )
    except Exception as e:
        shutil.rmtree(pptx_werkmap, ignore_errors=True)
        flash(f"Er ging iets mis bij het samenstellen: {e}")
        app.logger.error(traceback.format_exc())
        if werkmap:
            shutil.rmtree(werkmap, ignore_errors=True)
        try:
            schrijf_dienstverslag(
                uitvoer_map=OUTPUT_DIR,
                bestandsnaam=bestandsnaam,
                liturgie=liturgie,
                kerk=kerk,
                datum=datum,
                aanvangstijd=aanvangstijd,
                voorganger=voorganger,
                organist=organist,
                bijbel_vertaling=bijbel_vertaling,
                weerklank_afbeeldingen=weerklank_afbeeldingen,
                weerklank_fouten=weerklank_fouten,
                bevindingen=bevindingen,
                mislukt_reden=str(e),
            )
        except Exception:
            pass
        return redirect(url_for("index"))

    shutil.rmtree(pptx_werkmap, ignore_errors=True)
    if werkmap:
        shutil.rmtree(werkmap, ignore_errors=True)

    # Monitorverslag naar schijf: een map per dienst met de datum in de
    # naam, daarin de controlescherm-info + wat er goed ging of fout liep.
    try:
        schrijf_dienstverslag(
            uitvoer_map=OUTPUT_DIR,
            bestandsnaam=bestandsnaam,
            liturgie=liturgie,
            kerk=kerk,
            datum=datum,
            aanvangstijd=aanvangstijd,
            voorganger=voorganger,
            organist=organist,
            bijbel_vertaling=bijbel_vertaling,
            voorzang_opgezocht=voorzang_opgezocht,
            collecte_opgezocht=collecte_opgezocht,
            collectedoelen=collectedoelen,
            weerklank_afbeeldingen=weerklank_afbeeldingen,
            weerklank_fouten=weerklank_fouten,
            bevindingen=bevindingen,
        )
    except Exception as e:
        app.logger.error("Dienstverslag schrijven mislukt: %s", e)

    # Resultaatpagina: het downloaden start vanzelf, en de aandachtspunten
    # (wat er niet gevonden of afgeleid is) staan hier zichtbaar -- niet meer
    # alleen verstopt in het dienstverslag op schijf.
    return render_template("klaar.html", bestandsnaam=bestandsnaam,
                           bevindingen=bevindingen,
                           psalmen=([regel_label(r) for r in liturgie.get("muziekmomenten", []) if r]
                                    if profile_cfg.get("kind") == "voorbeeld" else
                                    [regel_label(r) if r.get("bundel") == "OB" else
                                     regel_label(r) + " (Weerklank/lied)"
                                     for r in (liturgie.get("secties") or {}).values()]),
                           voorganger=voorganger, datum=datum)


@app.route("/download/<path:bestandsnaam>")
def download(bestandsnaam):
    naam = os.path.basename(bestandsnaam)
    pad = os.path.join(OUTPUT_DIR, naam)
    if not naam.lower().endswith(".pptx") or not os.path.isfile(pad):
        flash("Dat bestand bestaat niet (meer). Maak de presentatie opnieuw.")
        return redirect(url_for("index"))
    return send_file(pad, as_attachment=True, download_name=naam)

if __name__ == "__main__":
    # Debug staat UIT (de Werkzeug-debugger op 0.0.0.0 laat iedereen op het netwerk
    # code uitvoeren). Voor ontwikkelen: FLASK_DEBUG=1 python app.py
    # Voor de homeserver: gunicorn -w 2 -b 0.0.0.0:5000 app:app
    app.run(host=os.environ.get("HOST", "0.0.0.0"), port=int(os.environ.get("PORT", "5000")),
            debug=os.environ.get("FLASK_DEBUG") == "1")
