"""
assemble.py
============
Voegt alles samen: een geparste liturgie (uit liturgie_parser.py), de
psalm-bibliotheek-index (psalm_index.json) en het kerksjabloon
(template_map.json + de .pptx zelf) tot een complete, kant-en-klare
dienst-PowerPoint.

Dit is de generieke versie van wat we in de chat 2x met de hand deden voor
de morgendienst van ds. Mulder en de avonddienst van ds. Van Oostende --
zelfde stappen, nu automatisch en herbruikbaar voor élke volgende liturgie
die bij dit ene sjabloon past.

GEBRUIK (zie ook app.py voor de webversie):

    from assemble import assemble_dienst

    assemble_dienst(
        liturgie=parse_liturgie("liturgie.docx"),
        kerk="MK",                      # GK of MK -- niet uit de liturgie zelf te halen
        organist="Jan Dekker",          # idem
        voorganger_override=None,       # alleen invullen als de liturgie het zelf niet zegt
        aanvangstijd="18:00",           # idem
        collectedoelen=["Evangelisatie", "Kerkrentmeesterlijk Beheer"],  # optioneel, uit collecterooster
        bijbel_vertaling="HSV",         # HSV | NBV | NBV21 -- voor de schriftlezingen
        weerklank_afbeeldingen=None,    # {sectie_naam: [{'pad','breedte','hoogte'}]} uit weerklank.verwerk_zip
        output_path="resultaat.pptx",
    )

BELANGRIJKE BEPERKINGEN (bewust, zie de conversatie voor waarom):
- Secties met een Weerklank-lied/psalm ('bundel' == 'Wk') worden NIET uit de
  eigen psalmbibliotheek gevuld: die liederen komen uit weerklank.liedbundels.nu.
  In de webapp kun je per Weerklank-lied het liedbundel-zipje uploaden; de app
  knipt die png's bij (port van de losse kerkdia-app, zie weerklank.py) en
  plakt de dia's automatisch achter de titel-dia van dat lied. Zonder zipje
  blijft er een titel-only placeholder staan (zoals voorheen), waarachter je
  de dia's zelf plakt. En áls een psalmnummer (toch) niet in de
  psalmbibliotheek zit (Psalm 1-150), blijft er ook alleen een titel-dia
  staan, met een melding in het dienstverslag -- nooit een crash.
- Schriftlezingen worden automatisch gevuld met echte bijbeltekst uit
  data/bijbeltekst/ (HSV/NBV/NBV21, kies de vertaling in de webapp). Dit
  is de opmaak van de losse kerkdia-txt-app: een titeldia met de
  verwijzing, daarna tekstdia's met maximaal 8 regels en de versnummers
  in superscript. Voor een schriftlezing moeten het hele verzen zijn;
  lukt het de verwijzing niet op te lossen (onbekend boek, deels verzen
  zoals 'Ezechiël 47: 9b+12', of ontbrekende data), dan komt er net als
  vroeger een lege placeholder-dia waar je de tekst zelf in plakt.
  De KERNTEKST-dia's vullen wél ook zulke deels-verwijzingen: de hele
  verzen plus de a/b-delen (gesplitst op de zinsgrens dichtst bij het
  midden van het vers; best effort).
- Een sectie die in de liturgie ontbreekt (bv. geen voorzang opgegeven)
  wordt gewoon uit het sjabloon verwijderd, niet geraden.
"""
import json, os, re, sys

sys.path.insert(0, os.path.dirname(__file__))
from pptx_surgery import PptxWorkspace
from bijbeltekst import (STANDAARD_VERTALING,
                         verwijzing_parsen, verwijzing_ophalen,
                         verwijzing_titeltekst, tekst_paginas, regels_naar_xml)
from weerklank import DIA_BREEDTE, DIA_HOOGTE, plaatsing


# Kerntekst-dia: de tekst komt in een vaste tekstvak-breedte (dezelfde als de
# bijbeltekstdia's). Een lange kerntekst wordt niet meer geweigerd maar
# automatisch kleiner gezet tot hij past (vloer = 18 pt). Deze breedte is in
# 'schaal-100px'-eenheden t.o.v. de echte beschikbare breedte van het
# tekstvak op de kerntekst-dia (ruim 11 inch bij 96 dpi, minus insets).
_KERNTekst_BREEDTE_PX96 = 1072     # beschikbare regelbreedte van Tekstvak 4
_KERNTekst_HOOGTE_INCH = 4.2       # ruimte voor de tekstregels onder de kop
_KERNTekst_FORMATEN = [2800, 2400, 2100, 1800]  # hondersten pt, aflopend


def _kerntekst_regels(tekst, sz):
    """Hoeveel regels wordt `tekst` in de kerntekst-tekstvak bij lettergrootte
    `sz` (hondersten pt)? Meet de woorden met PIL op de echte Century Gothic,
    dezelfde meetmethode als bijbeltekst.py gebruikt voor de lezingen."""
    from PIL import ImageFont
    font = ImageFont.truetype(os.path.join(os.path.dirname(__file__),
                                           "data", "fonts", "GOTHIC.TTF"), 100)
    pt = sz / 100.0
    schaal = (pt * 96 / 72) / 100.0          # 100px-meting -> werkelijke px
    limiet = _KERNTekst_BREEDTE_PX96
    regels = 1
    breedte = 0.0
    for woord in tekst.split():
        w = font.getlength(woord) * schaal
        if breedte + w > limiet and breedte > 0:
            regels += 1
            breedte = w
        else:
            breedte += w + font.getlength(" ") * schaal
    return regels


def _kerntekst_maat(tekst):
    """Kiest het grootste formaat (2800/2400/2100/1800) waarop de tekst in de
    kerntekst-dia past. Geeft (sz, regels) of (None, None) als zelfs 18 pt
    niet past (dan blijft de placeholder staan)."""
    for sz in _KERNTekst_FORMATEN:
        regels = _kerntekst_regels(tekst, sz)
        hoogte = regels * (sz / 100.0 / 72.0) * 1.25
        if hoogte <= _KERNTekst_HOOGTE_INCH and regels <= 14:
            return sz, regels
    return None, None



def _pas_lezingen_in(lezingen, plekken):
    """Het sjabloon heeft maar een vast aantal tekstregels voor schriftlezingen
    (overzichtsdia, orde van dienst). Zijn er meer lezingen dan plekken, dan
    worden de laatste samengevoegd op de laatste regel ('A; B') zodat er geen
    stilletjes een lezing wegvalt."""
    lezingen = [l for l in lezingen if l]
    if plekken <= 0:
        return []
    if len(lezingen) <= plekken:
        return lezingen
    return lezingen[:plekken - 1] + ["; ".join(lezingen[plekken - 1:])]


def _geldige_tijd(aanvangstijd):
    """'10:45' -> ('10', '45'); alles wat geen HH:MM is -> None."""
    m = re.match(r"^(\d{1,2}):(\d{2})$", (aanvangstijd or "").strip())
    if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
        return None
    return m.group(1).zfill(2), m.group(2)


def _meld(bevindingen, tekst):
    """Vangt een opmerking over het verloop van de samenstelling op: in de
    console (zoals voorheen) én -- als de aanroeper een lijst meegeeft -- in
    het dienstverslag dat na het genereren naar schijf wordt geschreven
    (zie dienstverslag.py)."""
    if bevindingen is not None:
        bevindingen.append(tekst)
    print(tekst)


def _load_json(path):
    return json.load(open(path, encoding="utf-8"))


def _psalm_entry(psalm_index, nummer, vers_nums=None):
    """Zoekt de index-entry voor een psalmnummer. Sommige psalmen (zoals
    Psalm 119) zijn opgesplitst over MEERDERE bestanden -- als vers_nums
    is opgegeven, wordt specifiek het bestand gekozen dat die verzen ook
    echt bevat, niet zomaar de eerst-gevonden entry."""
    kandidaten = [entry for key, entry in psalm_index.items()
                  if key.startswith(nummer + "::") and entry.get("verses")]
    if not kandidaten:
        return None
    if vers_nums:
        for entry in kandidaten:
            vmap = {v["vers"] for v in entry["verses"]}
            if all(vn in vmap for vn in vers_nums):
                return entry
    return kandidaten[0]


def _verse_image_paths(psalm_index, render_cache_dir, nummer, vers_nums, soort="Psalm"):
    entry = _psalm_entry(psalm_index, nummer, vers_nums)
    if entry is None:
        raise ValueError(f"{soort} {int(nummer)} niet (met hoog vertrouwen) gevonden in de index")
    safe = re.sub(r"[^\w]", "_", entry["file"])
    work_dir = os.path.join(render_cache_dir, safe)
    if not os.path.isdir(work_dir):
        # val terug op de simpelere sanitizing die eerder in de chat gebruikt is
        safe = entry["file"].replace("(", "").replace(")", "").replace(" ", "_")
        work_dir = os.path.join(render_cache_dir, safe)

    vmap = {v["vers"]: v["slides"] for v in entry["verses"]}
    paths = []
    for vn in vers_nums:
        if vn not in vmap:
            raise ValueError(f"Vers {vn} niet gevonden voor {soort} {int(nummer)}")
        for slide_num in vmap[vn]:
            match = [f for f in os.listdir(work_dir)
                     if f.endswith(".jpg") and int(re.search(r"(\d+)", f).group(1)) == slide_num]
            if not match:
                raise FileNotFoundError(f"dia {slide_num} niet gevonden in {work_dir}")
            paths.append(os.path.join(work_dir, match[0]))
    return paths


def _nummer_tekst(parsed_regel):
    nummer = str(parsed_regel.get("nummer", "0"))
    try:
        nummer = str(int(nummer))
    except ValueError:
        pass
    return nummer + str(parsed_regel.get("suffix", ""))


def _sectie_titel(parsed_regel):
    verzen_str = ", ".join(str(v) + ("her" if v in parsed_regel.get("herhaal", []) else "")
                            for v in parsed_regel.get("verzen", []))
    soort = parsed_regel.get("soort", "Lied")
    nummer = _nummer_tekst(parsed_regel)
    if parsed_regel.get("bundel") == "Wk":
        prefix = "Wk " if soort == "Psalm" else "Weerklank "
        return f"{prefix}{soort} {nummer}: {verzen_str}"
    return f"{soort} {nummer}: {verzen_str}"


def _vul_muziekmoment(ws, slot_config, parsed_regel, psalm_index,
                      render_cache_dir, *, gezang_index=None,
                      wk_plaatjes=None, bevindingen=None, label=None):
    """Vult één liedmoment in een ingevuld voorbeeldsjabloon.

    De titel-dia's en oude verzen liggen per dienstprofiel vast. Zo blijven
    alle vaste formulier- en avondmaalsdia's uit het voorbeeld behouden.
    """
    titel_dias = slot_config.get("titel_dias", [])
    oude_versdias = slot_config.get("oude_versdias", [])
    if not parsed_regel:
        ws.remove_slides(titel_dias + oude_versdias)
        return

    titel = _sectie_titel(parsed_regel)
    if label == "Gebed des Heeren":
        titel = "Gezang 5, Gebed des Heeren"
    pattern = (r"(?:Weerklank\s+|Wk\s+)?(?:Psalm|Ps\.?|Gezang|Gz\.?|"
               r"Lied|Ld)\s+\d{1,3}[A-Za-z]?(?::?\s*[^<]*)?")
    for dia in titel_dias:
        vervangen = ws.replace_text_pattern(dia, pattern, titel)
        if not vervangen:
            raise ValueError(f"Geen liedverwijzing gevonden op profiel-dia {dia}; "
                             "de oude titel is behouden om een verkeerde presentatie te voorkomen.")
    ws.remove_slides(oude_versdias)

    if parsed_regel.get("bundel") == "Wk":
        plaatjes = wk_plaatjes or []
        if not plaatjes:
            return
        layout = (ws.get_slide_layout(oude_versdias[0]) if oude_versdias else None) or "slideLayout4.xml"
        anchor = ws.get_rid_for_slide(titel_dias[-1])
        vorige_breedte = 0
        for info in plaatjes:
            links, boven, breedte, hoogte, vorige_breedte = plaatsing(
                DIA_BREEDTE, DIA_HOOGTE, info, vorige_breedte)
            anchor = ws.add_picture_slide_after(
                anchor, info["pad"], layout=layout,
                slide_width=breedte, slide_height=hoogte,
                left=links, top=boven)
        return

    soort = parsed_regel.get("soort", "Psalm")
    index = gezang_index if soort == "Gezang" else psalm_index
    verzen = parsed_regel.get("verzen", [])
    if index is None:
        _meld(bevindingen, f"Geen index voor {soort} beschikbaar; alleen de titel-dia is neergezet.")
        return
    herhaal = parsed_regel.get("herhaal", [])
    uit_te_voegen = [v for vers in verzen for v in ([vers, vers] if vers in herhaal else [vers])]
    dia_paden = []
    for vers in uit_te_voegen:
        try:
            dia_paden += _verse_image_paths(index, render_cache_dir,
                                            parsed_regel["nummer"], [vers], soort=soort)
        except (ValueError, FileNotFoundError) as e:
            _meld(bevindingen, f"{titel}: {e}; vers overgeslagen.")
    anchor = ws.get_rid_for_slide(titel_dias[-1])
    for img in dia_paden:
        anchor = ws.add_picture_slide_after(anchor, img)


def _vul_sectie(ws, sectie_config, parsed_regel, psalm_index, render_cache_dir,
                bevindingen=None, gezang_index=None):
    titel_dias = sectie_config["titel_dias"]
    oude_versdias = sectie_config["oude_versdias"]
    titel_tekst = _sectie_titel(parsed_regel)

    for t in titel_dias:
        ws.replace_all_text(t, [titel_tekst])

    ws.remove_slides(oude_versdias)

    # Alleen echte psalmen (Psalm 1-150) en de eigen gezangen (Gezang 1-12)
    # uit de eigen bibliotheek (bundel 'OB') worden automatisch ingevoegd. Weerklank-psalmen én alle liederen en
    # gezangen (bundel 'Wk': die komen uit weerklank.liedbundels.nu / het
    # Liedboek, niet uit onze psalmbibliotheek) krijgen hier géén eigen
    # dia's: alleen de titel-dia als placeholder. Worden er op het
    # controlescherm liedbundel-zipjes geüpload (zie weerklank.py), dan
    # komen die omgezette dia's in een latere stap automatisch achter deze
    # titel-dia terecht (zie _vul_weerklank_afbeeldingen).
    if parsed_regel["bundel"] != "OB":
        return

    # Gezang 1-12 komen uit hun eigen index (gezang_index.json), psalmen uit
    # psalm_index.json -- de nummers overlappen (Psalm 5 <> Gezang 5)
    soort = parsed_regel.get("soort", "Psalm")
    index = gezang_index if soort == "Gezang" else psalm_index
    if index is None:
        _meld(bevindingen, f"Geen index voor {soort} beschikbaar (data/gezang_index.json ontbreekt) "
                           "-- alleen een titel-placeholder-dia neergezet.")
        return

    # herhaling ("4her") betekent: dat vers 2x plakken
    verzen_met_herhaling = []
    for v in parsed_regel["verzen"]:
        verzen_met_herhaling.append(v)
        if v in parsed_regel["herhaal"]:
            verzen_met_herhaling.append(v)

    # Vers voor vers plakken. Ontbreekt een vers (of de hele psalm) in de
    # bibliotheek, dan worden de verzen die er wél zijn gewoon geplakt en komt
    # het ontbrekende vers als aandachtspunt in het verslag (zie het scherm na
    # het downloaden) -- de presentatie crasht er niet op.
    dia_paden = []
    ontbrekend = []
    for vn in verzen_met_herhaling:
        try:
            dia_paden += _verse_image_paths(index, render_cache_dir,
                                            parsed_regel["nummer"], [vn], soort=soort)
        except (ValueError, FileNotFoundError) as e:
            if vn not in ontbrekend:
                ontbrekend.append(vn)
            print(e)  # de samenvattende melding hieronder komt in het verslag
    if ontbrekend:
        _meld(bevindingen,
              f"{titel_tekst}: vers {', '.join(str(v) for v in ontbrekend)} staat niet in de "
              "bibliotheek en is overgeslagen"
              + ("; alleen de titel-dia is neergezet." if not dia_paden else "."))
        if not dia_paden:
            return

    anchor = ws.get_rid_for_slide(titel_dias[-1])
    for img in dia_paden:
        anchor = ws.add_picture_slide_after(anchor, img)


def _verwijder_sectie(ws, sectie_config):
    """Sectie ontbreekt in de liturgie -> titel + oude versdia's gewoon weg."""
    ws.remove_slides(sectie_config["titel_dias"] + sectie_config["oude_versdias"])


def _vul_weerklank_afbeeldingen(ws, liturgie, tmap, afbeeldingen, bevindingen=None):
    """Plakt de omgezette liedbundel-png's achter de titel-dia van het
    bijbehorende Weerklank-lied ('afbeeldingen': sectie_naam -> lijst van
    {'pad','breedte','hoogte'}, zie weerklank.verwerk_zip). Elke png wordt
    gecentreerd op een eigen dia geplaatst (zelfde breedte voor
    opeenvolgende verzen, zoals de kerkdia-app deed).
    """
    for sectie_naam, plaatjes in afbeeldingen.items():
        sectie_cfg = tmap.get("secties", {}).get(sectie_naam)
        parsed = liturgie.get("secties", {}).get(sectie_naam)
        if not sectie_cfg or not parsed or parsed.get("bundel") != "Wk":
            _meld(bevindingen, f"Weerklank-dia's voor '{sectie_naam}' overgeslagen "
                  "(geen Weerklank-plek in deze liturgie).")
            continue
        titel_dias = sectie_cfg.get("titel_dias") or []
        if not titel_dias:
            _meld(bevindingen, f"Weerklank-dia's voor '{sectie_naam}': geen titel-dia om achter te plakken.")
            continue
        # zelfde layout als de oude versdia's van deze plek (bv. de lichte
        # achtergrond waarop de liedbundel-dia's vroeger ook lagen)
        layout = "slideLayout4.xml"
        oude = sectie_cfg.get("oude_versdias") or []
        if oude:
            layout = ws.get_slide_layout(oude[0]) or layout
        anker = ws.get_rid_for_slide(titel_dias[-1])
        vorige_breedte = 0
        for info in plaatjes:
            links, boven, breedte, hoogte, vorige_breedte = plaatsing(
                DIA_BREEDTE, DIA_HOOGTE, info, vorige_breedte)
            anker = ws.add_picture_slide_after(
                anker, info["pad"], layout=layout,
                slide_width=breedte, slide_height=hoogte,
                left=links, top=boven)
        _meld(bevindingen, f"Weerklank '{sectie_naam}': {len(plaatjes)} lieddia's "
              "achter de titel-dia geplakt.")


def _vul_schriftlezing(ws, lezing_tekst, anker_rid, layout, vertaling, placeholder,
                       bevindingen=None):
    """Zet één schriftlezing neer: een titeldia met de verwijzing, gevolgd
    door de opgemaakte bijbeltekstdia's (versregels, zoals kerkdia-txt die
    maakte). Kan de verwijzing niet (volledig) opgelost worden -- onbekend
    boek, deels verzen zoals '9b+12', of de vertaling ontbreekt -- dan komt
    er een titeldia met de ruwe tekst plus één lege placeholder-dia.
    Retourneert het rId van de laatst toegevoegde dia."""
    verwijzing = verwijzing_parsen(lezing_tekst)
    verzen = verwijzing_ophalen(vertaling, verwijzing) if verwijzing else None

    if not verzen:
        _meld(bevindingen, f"Schriftlezing '{lezing_tekst}' niet automatisch "
              "gevuld (onbekend boek, deels verzen of tekst niet beschikbaar): "
              "titel-dia + lege tekst-dia.")
        rid = ws.add_lezingtitel_slide_after(anker_rid, (lezing_tekst or "").strip(), layout)
        return ws.add_bijbeltekstdia_after(rid, regels_naar_xml([[( "n", placeholder)]]), layout)

    rid = ws.add_lezingtitel_slide_after(anker_rid, verwijzing_titeltekst(verwijzing), layout)
    for pagina in tekst_paginas(verzen):
        rid = ws.add_bijbeltekstdia_after(rid, regels_naar_xml(pagina), layout)
    return rid


def assemble_dienst(liturgie, kerk, output_path,
                     organist=None, voorganger_override=None,
                     aanvangstijd=None, collectedoelen=None,
                     data_dir=None, work_dir="/tmp/pptx_werkmap",
                     bijbel_vertaling=STANDAARD_VERTALING,
                     weerklank_afbeeldingen=None,
                     bevindingen=None,
                     service_profile="standard"):
    data_dir = data_dir or os.path.join(os.path.dirname(__file__), "data")
    profiles_path = os.path.join(data_dir, "service_profiles.json")
    profiles = _load_json(profiles_path) if os.path.exists(profiles_path) else {}
    profile = profiles.get(service_profile) or profiles.get("standard") or {}
    # Eén zichtbare categorie Avondmaalsdienst; kies intern het juiste
    # voorbeeld voor de tweede ochtenddienst op basis van de ingevulde tijd.
    if profile.get("variants"):
        variant = "tweede_ochtend" if aanvangstijd == "11:00" else "ochtend"
        profile = {**profile, **profile["variants"][variant]}
    if profile.get("kind") == "voorbeeld":
        tmap = profile["template_map"]
        template_file = profile["template_file"]
    else:
        map_file = profile.get("template_map", "template_map.json")
        tmap = _load_json(os.path.join(data_dir, map_file))
        template_file = tmap["template_file"]
    psalm_index = _load_json(os.path.join(data_dir, "psalm_index.json"))
    gezang_pad = os.path.join(data_dir, "gezang_index.json")
    gezang_index = _load_json(gezang_pad) if os.path.exists(gezang_pad) else None
    render_cache_dir = os.path.join(data_dir, "render_cache")
    template_pptx = os.path.join(data_dir, template_file)

    ws = PptxWorkspace.unpack(template_pptx, work_dir)

    # -- liederen -------------------------------------------------------
    if "muziekmomenten" in tmap:
        momenten = liturgie.get("muziekmomenten", [])
        slots = tmap["muziekmomenten"]
        if len(momenten) > len(slots):
            raise ValueError(f"Deze dienstindeling heeft {len(slots)} liedplekken, maar de liturgie bevat "
                             f"{len(momenten)} liedmomenten. Er is niets stil overgeslagen.")
        for i, slot in enumerate(slots):
            parsed = momenten[i] if i < len(momenten) else None
            _vul_muziekmoment(
                ws, slot, parsed, psalm_index, render_cache_dir,
                gezang_index=gezang_index,
                wk_plaatjes=(weerklank_afbeeldingen or {}).get(str(i)),
                bevindingen=bevindingen,
                label=(parsed or {}).get("label"),
            )
    else:
        for sectie_naam, sectie_config in tmap["secties"].items():
            if sectie_naam.startswith("_"):
                continue
            parsed_regel = liturgie.get("secties", {}).get(sectie_naam)
            if parsed_regel:
                _vul_sectie(ws, sectie_config, parsed_regel, psalm_index, render_cache_dir,
                            bevindingen=bevindingen, gezang_index=gezang_index)
            else:
                _verwijder_sectie(ws, sectie_config)

    # -- Weerklank-liedbundel-dia's (geüploade zipjes) -------------------
    # gebeurt ná de psalmsecties: de titel-dia's van Weerklank-plekken
    # staan dan al klaar (zie _vul_sectie) en de png's komen er direct
    # achter te liggen, precies zoals bij de psalmvers-dia's.
    if weerklank_afbeeldingen and "muziekmomenten" not in tmap:
        _vul_weerklank_afbeeldingen(ws, liturgie, tmap, weerklank_afbeeldingen,
                                    bevindingen=bevindingen)

    # -- Wet van God / Geloofsbelijdenis ---------------------------------
    wet_cfg = tmap.get("wet_en_geloofsbelijdenis")
    if wet_cfg:
        is_ochtend = liturgie.get("type_dienst") == "ochtend"
        ws.set_slide_hidden(wet_cfg["wet_slide"], hidden=not is_ochtend)
        ws.set_slide_hidden(wet_cfg["geloofsbelijdenis_slide"], hidden=is_ochtend)

    # -- titeldia + welkomstdia ------------------------------------------
    voorganger = voorganger_override or liturgie.get("voorganger") or "(nog aan te vullen)"
    thema = liturgie.get("thema")  # None als er geen thema is opgegeven -- NIET met een placeholder vullen
    # Niet iedereen is al dominee -- een proponent (kandidaat) krijgt geen
    # "Ds."-titel. Alleen "Ds. " ervoor zetten als de naam nog geen eigen
    # titel bevat.
    heeft_al_titel = bool(re.match(r'^(ds\.?|prop\.?|dhr\.?|drs\.?|dr\.?|prof\.?|kand\.?)\s', voorganger, re.IGNORECASE))
    titel_prefix = "" if heeft_al_titel else "Ds. "

    titel_cfg = tmap["titel_dia"]
    # Geen (geldige) aanvangstijd? Dan laten we het 'Aanvang: ... uur'-stuk weg in
    # plaats van '00:00' of een crash (vroeger: 'half tien' -> IndexError).
    tijd = _geldige_tijd(aanvangstijd)
    if aanvangstijd and not tijd:
        _meld(bevindingen, f"Aanvangstijd '{aanvangstijd}' is geen geldige tijd (HH:MM): "
              "weggelaten op de titeldia.")
    ws.replace_all_text(titel_cfg["slide"], [
        f"Welkom in deze dienst vanuit de {'Maranathakerk' if kerk == 'MK' else 'Grote Kerk'}",
        f"Aanvang: {tijd[0]}" if tijd else "",
        f":{tijd[1]}" if tijd else "",
        " uur" if tijd else "", titel_prefix, voorganger, f"Thema: {thema} " if thema else "",
    ])
    foto_rel = tmap["titel_dia"]["orgelfoto_per_kerk"].get(kerk) or \
        next(iter(tmap["titel_dia"]["orgelfoto_per_kerk"].values()))
    foto_pad = os.path.join(data_dir, foto_rel)
    if os.path.exists(foto_pad):
        ws.replace_slide_image(titel_cfg["slide"], foto_pad)

    welkom_slide = tmap["welkom_dia"]["slide"]
    ws.replace_all_text(welkom_slide, [
        "Welkom in deze dienst", "Voorganger:\t", f"{titel_prefix}{voorganger}",
        "Organist:", f"\t      {organist or '(nog aan te vullen)'}",
    ])

    # -- collectedoelen (optioneel) ---------------------------------------
    # Vult specifiek de 3 waarden na "Eerste/Tweede/Derde Rondgang:" in,
    # zonder de titel of eventuele "(zie vlg dia)"-opmerkingen te raken
    # (zie fill_rondgang_collecte in pptx_surgery.py voor waarom dat met de
    # oude, simpele tekstvervanging niet goed ging).
    if collectedoelen:
        for slide_num in tmap["collecte_doel_dias"]:
            try:
                ws.fill_rondgang_collecte(slide_num, collectedoelen)
            except Exception as e:
                _meld(bevindingen, f"Collectedoelen invullen op dia {slide_num} mislukt: {e}")

    # -- orde van dienst overzicht -----------------------------------------
    # LET OP: de volgorde hier is NIET gewoon de chronologische dienst-volgorde.
    # De "orde van dienst"-tabel in het sjabloon bestaat uit 2 kolommen (links
    # = eerste helft van de dienst, rechts = tweede helft), maar de tekst-runs
    # in de XML staan PER RIJ na elkaar (links-rij1, rechts-rij1, links-rij2,
    # rechts-rij2, ...). Om dus kolomsgewijs (boven-onder-links, dan
    # boven-onder-rechts) te lezen, moet de content in DEZE interleaved
    # volgorde aangeleverd worden. Geverifieerd tegen 2 verschillende echte
    # sjablonen -- beide gebruiken precies deze indeling.
    sectie_volgorde = ["voorzang", "voor_prediking", "opening", "na_prediking", "na_wet", "slot"]
    orde_regels = ["Orde van dienst"]
    for naam in sectie_volgorde:
        regel = liturgie.get("secties", {}).get(naam)
        if regel:
            soort_kort = {"Psalm": "Ps", "Lied": "Ld", "Gezang": "Gz"}.get(regel["soort"], "Ld")
            verzen_str = ", ".join(str(v) for v in regel["verzen"])
            orde_regels += [soort_kort, f"{int(regel['nummer'])}:", verzen_str]
        else:
            orde_regels += ["", "", ""]
    alle_lezingen = liturgie.get("schriftlezingen", [])
    for slide_num in tmap["orde_van_dienst_dias"]:
        try:
            if tmap.get("orde_van_dienst_mode") == "lezingen":
                refs = alle_lezingen
                ws.replace_reference_textbox(
                    slide_num, _pas_lezingen_in(refs, 2))
                continue
            plekken = ws.count_text_runs(slide_num) - len(orde_regels)
            ws.replace_all_text(slide_num, orde_regels + _pas_lezingen_in(alle_lezingen, plekken))
        except Exception as e:
            _meld(bevindingen, f"Orde van dienst op dia {slide_num} invullen mislukt: {e}")

    # -- schriftlezingen + kerntekst -----------------------------------------
    sl_cfg = tmap["schriftlezing"]
    lezingen = liturgie.get("schriftlezingen", [])
    placeholder = tmap["bijbeltekst_placeholder_tekst"]

    overzicht_plekken = ws.count_text_runs(sl_cfg["overzicht_dia"]) - 1
    if len(lezingen) > overzicht_plekken:
        _meld(bevindingen, f"Het schriftlezing-overzicht heeft maar {overzicht_plekken} regels: "
              "de laatste lezingen staan samen op de laatste regel (de lezingen zelf hebben wel "
              "hun eigen dia's).")
    ws.replace_all_text(sl_cfg["overzicht_dia"], ["Schriftlezingen"] + _pas_lezingen_in(lezingen, overzicht_plekken))

    # Referenties zonder bijbeltekst, zoals een catechismus-zondag, hebben in
    # sommige dienstvoorbeelden een eigen vaste dia. Houd ze buiten de
    # automatisch opgebouwde Bijbeltekstblokken.
    referentie_only = sl_cfg.get("referentie_only_dias", [])
    bijbel_lezingen = [l for l in lezingen if verwijzing_parsen(l)]
    losse_referenties = [l for l in lezingen if not verwijzing_parsen(l)]
    for index, slide_num in enumerate(referentie_only):
        if index < len(losse_referenties):
            ws.replace_all_text(slide_num, [losse_referenties[index]])
            ws.set_slide_hidden(slide_num, hidden=False)
        else:
            ws.remove_slides([slide_num])
    if len(losse_referenties) > len(referentie_only):
        _meld(bevindingen, f"Er zijn {len(losse_referenties)} niet-Bijbelverwijzingen, maar het sjabloon heeft "
              f"maar {len(referentie_only)} vaste referentiedia('s).")

    # oude lezing-dia's uit het sjabloon (van de vorige dienst) weg
    for slot in sl_cfg["lezingen"]:
        te_verwijderen = list(slot.get("oude_tekstdias", []))
        if slot.get("titel_dia") and slot["titel_dia"] not in te_verwijderen:
            te_verwijderen.append(slot["titel_dia"])
        ws.remove_slides(te_verwijderen)

    # dan per schriftlezing een blok (titel + bijbeltekstdia's) direct na
    # het overzicht plakken, zodat meerdere lezingen mooi na elkaar staan
    layout = None
    if sl_cfg.get("lezingen"):
        eerste_oude = sl_cfg["lezingen"][0].get("oude_tekstdias") or [None]
        layout = ws.get_slide_layout(eerste_oude[0]) if eerste_oude[0] else None
    anker_rid = ws.get_rid_for_slide(sl_cfg["overzicht_dia"])
    for lezing in bijbel_lezingen if referentie_only else lezingen:
        try:
            anker_rid = _vul_schriftlezing(ws, lezing, anker_rid, layout,
                                            bijbel_vertaling, placeholder,
                                            bevindingen=bevindingen)
        except Exception as e:
            _meld(bevindingen, f"Schriftlezing '{lezing}' kon niet automatisch "
                  f"gevuld worden ({e}); lege placeholder-dia.")
            try:
                rid = ws.add_lezingtitel_slide_after(anker_rid, (lezing or "").strip(), layout)
                anker_rid = ws.add_bijbeltekstdia_after(rid, regels_naar_xml([[( "n", placeholder)]]), layout)
            except Exception:
                pass

    # -- kerntekst -------------------------------------------------------------
    kerntekst = liturgie.get("kerntekst")
    if kerntekst:
        # Kan de kerntekst automatisch gevuld worden? Alleen hele verzen
        # (deels verzen zoals '9b+12' kunnen niet uit de JSON gehaald
        # worden). Een lange kerntekst wordt niet geweigerd maar automatisch
        # kleiner gezet (28 -> 24 -> 21 -> 18 pt) tot hij in het tekstvak
        # past; past ook 18 pt niet, dan blijft de
        # '[ Bijbeltekst hier plakken ]'-placeholder staan.
        kerntekst_tekst = None
        kerntekst_maat = None
        verwijzing = verwijzing_parsen(kerntekst)
        # met_delen=True: ook deels-verwijzingen zoals 'Ezechiël 47: 9b+12'
        # worden opgehaald (de a/b-delen op zinsgrenzen gesplitst).
        verzen = (verwijzing_ophalen(bijbel_vertaling, verwijzing, met_delen=True)
                  if verwijzing else None)
        if verzen:
            stukken = []
            for hfdst in sorted(verzen, key=int):
                for versnr in sorted(verzen[hfdst],
                                     key=lambda k: (int(re.match(r"\d+", k).group(0)), k)):
                    stukken.append(f"{versnr} {verzen[hfdst][versnr].strip()}")
            kandidaat = " ".join(stukken)
            kerntekst_maat, _ = _kerntekst_maat(kandidaat)
            if kerntekst_maat:
                kerntekst_tekst = kandidaat

        if verzen and not kerntekst_maat:
            _meld(bevindingen, f"Kerntekst '{kerntekst}': de tekst past niet in "
                  "het tekstvak (zelfs niet op 18 pt); de placeholder-dia "
                  "blijft staan.")
        elif not verzen:
            _meld(bevindingen, f"Kerntekst '{kerntekst}' kon niet automatisch "
                  "gevuld worden (onbekend boek of tekst niet beschikbaar); "
                  "de placeholder-dia blijft staan.")

        for slide_num in sl_cfg["kerntekst_dias"]:
            try:
                ws.set_slide_hidden(slide_num, hidden=False)
                ws.replace_all_text(slide_num, ["Kerntekst", " ", placeholder])
                # referentie onderaan (run met sz="4400" op deze dia's)
                ws.set_first_run_text_with(slide_num, 'sz="4400"', f" {kerntekst}")
                if kerntekst_tekst and kerntekst_maat:
                    ws.set_first_run_text_with(slide_num, 'sz="2800"',
                                               kerntekst_tekst,
                                               nieuwe_sz=str(kerntekst_maat))
            except Exception as e:
                _meld(bevindingen, f"Kerntekst-dia {slide_num} invullen mislukt: {e}")
        ws.set_title_text(sl_cfg["verkondiging_referentie_dia"], f" {kerntekst}", run_index=1)
    else:
        # géén kerntekst in de liturgie: de vaste kerntekst-dia's uit het
        # sjabloon leegmaken, zodat er geen oude tekst van een vorige dienst
        # in de presentatie blijft staan.
        for slide_num in sl_cfg["kerntekst_dias"]:
            try:
                ws.set_slide_hidden(slide_num, hidden=True)
                ws.replace_all_text(slide_num, ["Kerntekst", " ", ""])
                ws.set_first_run_text_with(slide_num, 'sz="4400"', "")
            except Exception:
                pass
    for slide_num in sl_cfg["thema_dias"][1:]:  # dia 1 (titeldia) al gedaan hierboven
        try:
            if thema:
                ws.set_title_text(slide_num, f"Thema: {thema}")
                ws.set_slide_hidden(slide_num, hidden=False)
            else:
                ws.set_slide_hidden(slide_num, hidden=True)
        except Exception:
            pass

    ws.repack(output_path)
    return output_path


if __name__ == "__main__":
    from liturgie_parser import parse_liturgie

    if len(sys.argv) < 4:
        print("gebruik: python assemble.py <liturgie.docx> <GK|MK> <output.pptx>")
        sys.exit(1)

    liturgie_pad, kerk, output_pad = sys.argv[1], sys.argv[2], sys.argv[3]
    liturgie = parse_liturgie(liturgie_pad)
    print(json.dumps(liturgie, indent=2, ensure_ascii=False))

    assemble_dienst(liturgie=liturgie, kerk=kerk, output_path=output_pad,
                     aanvangstijd="18:00" if liturgie.get("type_dienst") == "avond" else "09:30")
    print(f"\nKlaar: {output_pad}")
