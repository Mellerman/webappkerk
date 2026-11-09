"""
pptx_surgery.py
================
Laagniveau hulpfuncties om een BESTAAND .pptx-bestand te bewerken door
rechtstreeks in de "uitgepakte" XML te werken (een .pptx is gewoon een zip-
bestand met XML-bestanden erin).

WAAROM NIET GEWOON python-pptx GEBRUIKEN?
python-pptx kan prima NIEUWE dia's opbouwen met tekst/vormen, maar heeft geen
ingebouwde functie om een dia te VERWIJDEREN of om een dia uit een ANDER
bestand te kopiëren. Voor dit project (dia's uit 150 losse bronbestanden in
één kerksjabloon plakken) is direct met de XML werken daarom simpeler en
betrouwbaarder dan het object-model heen-en-weer proberen te forceren.

BASISPRINCIPE VAN EEN .PPTX-BESTAND:
- Het is een zip met o.a.:
    ppt/presentation.xml            <- lijst van alle dia's (sldIdLst), IN VOLGORDE
    ppt/_rels/presentation.xml.rels <- koppelt elke sldId aan een slideN.xml bestand
    ppt/slides/slideN.xml           <- inhoud van dia N
    ppt/slides/_rels/slideN.xml.rels<- welke media/layout die dia gebruikt
    ppt/media/                      <- alle afbeeldingen
    [Content_Types].xml             <- registreert welke bestandstypen er zijn

GEBRUIK:
    1. Pak het .pptx-bestand uit met unpack(), dat geeft je een working directory.
    2. Gebruik de functies hieronder om te bewerken.
    3. Pak het weer in met repack().

VOORBEELD (zie ook fill_liturgie_template.py voor een compleet voorbeeld):

    from pptx_surgery import PptxWorkspace

    ws = PptxWorkspace.unpack("sjabloon.pptx", "/tmp/werkmap")
    ws.set_title_text(slide_num=14, new_text="Psalm 145: 1")
    ws.remove_slides([15, 16, 17, 18, 19, 20])
    anchor = ws.get_rid_for_slide(14)
    for img in ["/pad/naar/vers1_dia1.jpg", "/pad/naar/vers1_dia2.jpg"]:
        anchor = ws.add_picture_slide_after(anchor, img)
    ws.repack("resultaat.pptx")
"""
from xml.sax.saxutils import escape as _esc
import os, re, shutil, zipfile


class PptxWorkspace:
    """Wrapper rond een uitgepakt .pptx-bestand met bewerkingsfuncties."""

    def __init__(self, work_dir):
        self.dir = work_dir

    # -- in- en uitpakken --------------------------------------------------

    @classmethod
    def unpack(cls, pptx_path, work_dir):
        if os.path.exists(work_dir):
            shutil.rmtree(work_dir)
        os.makedirs(work_dir)
        with zipfile.ZipFile(pptx_path) as z:
            z.extractall(work_dir)
        return cls(work_dir)

    def repack(self, out_path):
        """Zipt de working directory weer in tot een geldig .pptx-bestand.
        BELANGRIJK: [Content_Types].xml moet als eerste in de zip staan en
        NIET gecomprimeerd raken op een manier die de structuur breekt --
        gewone zip -r werkt hiervoor prima."""
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(self.dir):
                for f in files:
                    full = os.path.join(root, f)
                    arcname = os.path.relpath(full, self.dir)
                    zf.write(full, arcname)
        return out_path

    # -- interne paden --------------------------------------------------

    def _p(self, *parts):
        return os.path.join(self.dir, *parts)

    def _read(self, *parts):
        return open(self._p(*parts), encoding="utf-8").read()

    def _write(self, content, *parts):
        open(self._p(*parts), "w", encoding="utf-8").write(content)

    # -- opzoeken -----------------------------------------------------------

    def get_rid_for_slide(self, slide_num):
        """Geeft het relationship-id (bv. 'rId19') terug dat bij
        ppt/slides/slideN.xml hoort volgens presentation.xml.rels."""
        rels = self._read("ppt/_rels/presentation.xml.rels")
        m = re.search(rf'Id="(rId\d+)"[^>]*Target="slides/slide{slide_num}\.xml"', rels)
        if not m:
            raise ValueError(f"slide{slide_num}.xml niet gevonden in presentation.xml.rels")
        return m.group(1)

    def _next_slide_num(self):
        nums = [int(re.search(r"slide(\d+)\.xml$", f).group(1))
                for f in os.listdir(self._p("ppt/slides")) if re.match(r"slide\d+\.xml$", f)]
        return max(nums) + 1

    def _next_media_num(self):
        nums = [int(re.match(r"image(\d+)\.\w+$", f).group(1))
                for f in os.listdir(self._p("ppt/media")) if re.match(r"image(\d+)\.\w+$", f)]
        return (max(nums) + 1) if nums else 1

    def _next_pres_rid(self):
        rels = self._read("ppt/_rels/presentation.xml.rels")
        ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
        return max(ids) + 1

    def _next_sldid(self):
        pres = self._read("ppt/presentation.xml")
        ids = [int(m) for m in re.findall(r'<p:sldId id="(\d+)"', pres)]
        return max(ids) + 1

    # -- bewerkingen ----------------------------------------------------

    def set_title_text(self, slide_num, new_text, run_index=0):
        """Vervangt de tekst van de N-de <a:t>-run op een dia (standaard de
        eerste). Handig voor titeldia's als 'Psalm 36: 2, 3' -> 'Psalm 145: 1'.
        Voor dia's met meerdere tekstvakken: geef alle nieuwe teksten als lijst
        aan replace_all_text() in plaats hiervan."""
        path = ("ppt/slides", f"slide{slide_num}.xml")
        xml = self._read(*path)
        positions = list(re.finditer(r"<a:t>[^<]*</a:t>", xml))
        if not positions:
            raise IndexError(f"Geen tekstrun gevonden op dia {slide_num}")
        # Sommige sjabloonvarianten hebben maar één tekstrun in de titel.
        # Gebruik in dat geval de eerste run in plaats van de hele dienst te
        # laten mislukken door een index die alleen bij andere sjablonen klopt.
        if run_index >= len(positions):
            run_index = 0
        m = positions[run_index]
        xml = xml[:m.start()] + f"<a:t>{_esc(new_text)}</a:t>" + xml[m.end():]
        self._write(xml, *path)

    def replace_all_text(self, slide_num, new_texts):
        """Vervangt ALLE <a:t>-tekstruns op een dia, in volgorde, door de
        gegeven lijst. Korter dan new_texts? De rest wordt leeggemaakt.
        Gebruik dit voor dia's met meerdere tekstvakken/tabellen, zoals een
        'orde van dienst'-overzicht."""
        path = ("ppt/slides", f"slide{slide_num}.xml")
        xml = self._read(*path)
        it = iter(new_texts)
        def repl(_):
            try:
                return f"<a:t>{_esc(next(it))}</a:t>"
            except StopIteration:
                return "<a:t></a:t>"
        xml = re.sub(r"<a:t>[^<]*</a:t>", repl, xml)
        self._write(xml, *path)

    def replace_reference_textbox(self, slide_num, new_texts, heading="Orde van dienst"):
        """Vervangt regels in het verwijzingstekstvak zonder tekstvakken of
        paragrafen te laten verschuiven. De meegeleverde dienstsjablonen hebben
        een aparte titel en een apart tekstvak met één lezing per regel."""
        from html import unescape
        path = ("ppt/slides", f"slide{slide_num}.xml")
        xml = self._read(*path)
        shapes = list(re.finditer(r"<p:sp(?:\s[^>]*)?>.*?</p:sp>", xml, re.S))
        target = None
        for shape in shapes:
            text = " ".join(unescape(t) for t in re.findall(r"<a:t>([^<]*)</a:t>", shape.group(0)))
            if text.strip() and heading.casefold() not in text.casefold():
                target = shape
                break
        if target is None:
            raise ValueError(f"Geen verwijzingstekstvak gevonden op dia {slide_num}")
        body = target.group(0)
        paragraphs = list(re.finditer(r"<a:p(?:\s[^>]*)?>.*?</a:p>", body, re.S))
        if not paragraphs:
            raise ValueError(f"Geen tekstregels gevonden op dia {slide_num}")
        values = list(new_texts)
        updated = body
        for pi, paragraph in reversed(list(enumerate(paragraphs))):
            value = values[pi] if pi < len(values) else ""
            para = paragraph.group(0)
            runs = list(re.finditer(r"<a:t>([^<]*)</a:t>", para))
            if not runs:
                continue
            for ri, run in reversed(list(enumerate(runs))):
                replacement = _esc(value) if ri == 0 else ""
                para = para[:run.start(1)] + replacement + para[run.end(1):]
            updated = updated[:paragraph.start()] + para + updated[paragraph.end():]
        xml = xml[:target.start()] + updated + xml[target.end():]
        self._write(xml, *path)

    def replace_text_pattern(self, slide_num, pattern, new_text, count=1):
        """Vervangt tekst in één PowerPoint-paragraaf, ook als de oude tekst
        over meerdere opgemaakte a:t-runs verdeeld is. Andere tekst op dezelfde
        dia (bijv. 'Bij voorbereiding tafel' of 'Toezingen ...') blijft staan."""
        from html import unescape

        path = ("ppt/slides", f"slide{slide_num}.xml")
        xml = self._read(*path)
        gewijzigd = 0
        edits = []
        for para in re.finditer(r"<a:p(?:\s[^>]*)?>.*?</a:p>", xml, re.S):
            runs = list(re.finditer(r"<a:t>([^<]*)</a:t>", para.group(0)))
            texts = [unescape(r.group(1)) for r in runs]
            combined = "".join(texts)
            match = re.search(pattern, combined, re.IGNORECASE)
            if not match:
                continue
            starts = []
            offset = 0
            for text in texts:
                starts.append(offset)
                offset += len(text)
            first = next((i for i, t in enumerate(texts)
                          if starts[i] + len(t) > match.start()), None)
            last = next((i for i in range(len(texts) - 1, -1, -1)
                         if starts[i] < match.end()), None)
            if first is None or last is None:
                continue
            prefix = texts[first][:max(0, match.start() - starts[first])]
            suffix = texts[last][max(0, match.end() - starts[last]):]
            replacements = [""] * len(texts)
            if first == last:
                replacements[first] = prefix + new_text + suffix
            else:
                replacements[first] = prefix + new_text
                replacements[last] = suffix
            for i in range(first, last + 1):
                run = runs[i]
                edits.append((para.start() + run.start(1), para.start() + run.end(1),
                              _esc(replacements[i])))
            gewijzigd += 1
            if gewijzigd >= count:
                break
        for start, end, replacement in reversed(edits):
            xml = xml[:start] + replacement + xml[end:]
        if gewijzigd:
            self._write(xml, *path)
        return gewijzigd

    def count_text_runs(self, slide_num):
        """Aantal <a:t>-tekstruns op een dia (zo weet assemble hoeveel regels er
        in bv. het schriftlezing-overzicht beschikbaar zijn)."""
        xml = self._read("ppt/slides", f"slide{slide_num}.xml")
        return len(re.findall(r"<a:t>[^<]*</a:t>", xml))

    def remove_slides(self, slide_nums):
        """Haalt dia's uit de zichtbare volgorde (sldIdLst). De onderliggende
        bestanden blijven fysiek bestaan (orphan parts) -- dat is onschadelijk
        en veel sneller/veiliger dan alles netjes opruimen."""
        pres = self._read("ppt/presentation.xml")
        for n in slide_nums:
            rid = self.get_rid_for_slide(n)
            pres = re.sub(rf'<p:sldId id="\d+" r:id="{rid}"/>', "", pres)
        self._write(pres, "ppt/presentation.xml")

    def set_slide_hidden(self, slide_num, hidden=True):
        """Verbergt of toont een dia (show="0" attribuut), zonder 'm te
        verwijderen. Gebruikt de kerk bijvoorbeeld voor Wet-van-God /
        Geloofsbelijdenis, die om en om zichtbaar zijn."""
        path = ("ppt/slides", f"slide{slide_num}.xml")
        xml = self._read(*path)
        m = re.search(r"<p:sld\b[^>]*>", xml)
        has_show0 = 'show="0"' in m.group(0) if m else False
        if hidden and not has_show0:
            xml = xml.replace("<p:sld ", '<p:sld show="0" ', 1)
        elif not hidden and has_show0:
            xml = xml.replace(' show="0"', "", 1)
        self._write(xml, *path)

    _PIC_SLIDE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''

    _PIC_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{media}"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/{layout}"/></Relationships>'''

    def add_picture_slide_after(self, after_rid, img_path, layout="slideLayout4.xml",
                                 slide_width=12192000, slide_height=6858000,
                                 left=0, top=0):
        """Voegt een nieuwe dia toe die bestaat uit precies 1 afbeelding
        (zoals de psalm-versdia's full-bleed, of de liedbundel-png's op hun
        eigen plek), direct na de dia met relationship-id `after_rid`.
        Standaard vult de afbeelding de hele dia (left/top 0, breedte/hoogte
        = diaformaat); voor een plaatje op maat geef je left/top en
        slide_width/slide_height zelf mee. Retourneert het nieuwe rId, zodat
        je opeenvolgende afbeeldingen na elkaar kunt plakken door steeds het
        laatst-teruggegeven rId als nieuwe 'after_rid' te gebruiken.
        """
        n = self._next_slide_num()
        media_num = self._next_media_num()
        ext = os.path.splitext(img_path)[1]
        media_name = f"image{media_num}{ext}"
        os.makedirs(self._p("ppt/media"), exist_ok=True)
        shutil.copy(img_path, self._p("ppt/media", media_name))

        self._write(self._PIC_SLIDE_XML.format(left=left, top=top,
                                               width=slide_width, height=slide_height),
                    "ppt/slides", f"slide{n}.xml")
        os.makedirs(self._p("ppt/slides/_rels"), exist_ok=True)
        self._write(self._PIC_RELS_XML.format(media=media_name, layout=layout),
                    "ppt/slides/_rels", f"slide{n}.xml.rels")

        self._ensure_jpg_registered()
        return self._register_slide(n, after_rid)

    def _ensure_jpg_registered(self):
        """.jpg wordt niet altijd standaard als content-type geregistreerd
        (vaak alleen .jpeg) -- dit voegt 'm toe als dat nog niet is gebeurd."""
        path = ("[Content_Types].xml",)
        xml = self._read(*path)
        if 'Extension="jpg"' not in xml:
            xml = xml.replace(
                '<Default Extension="jpeg" ContentType="image/jpeg"/>',
                '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                '<Default Extension="jpg" ContentType="image/jpeg"/>')
            self._write(xml, *path)

    def _register_slide(self, slide_num, after_rid):
        """Interne functie: registreert een nieuw slideN.xml bestand overal
        waar dat moet (rels, content-types, sldIdLst) en zet 'm direct na de
        dia met `after_rid` in de zichtbare volgorde."""
        rid = f"rId{self._next_pres_rid()}"

        rels_path = ("ppt/_rels/presentation.xml.rels",)
        rels = self._read(*rels_path)
        new_rel = (f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                   f'officeDocument/2006/relationships/slide" Target="slides/slide{slide_num}.xml"/>')
        self._write(rels.replace("</Relationships>", new_rel + "</Relationships>"), *rels_path)

        ct_path = ("[Content_Types].xml",)
        ct = self._read(*ct_path)
        new_override = (f'<Override PartName="/ppt/slides/slide{slide_num}.xml" '
                         f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
        self._write(ct.replace("</Types>", new_override + "</Types>"), *ct_path)

        pres_path = ("ppt/presentation.xml",)
        pres = self._read(*pres_path)
        sldid = self._next_sldid()
        new_entry = f'<p:sldId id="{sldid}" r:id="{rid}"/>'
        marker = f'r:id="{after_rid}"/>'
        pos = pres.index(marker) + len(marker)
        self._write(pres[:pos] + new_entry + pres[pos:], *pres_path)

        return rid

    def replace_slide_image(self, slide_num, new_image_path):
        """Vervangt de afbeelding die op een dia wordt gebruikt (bv. de
        orgelfoto op de titeldia) door een andere, zonder de XML/rels aan
        te hoeven passen -- we overschrijven gewoon het media-bestand zelf
        op dezelfde plek. Werkt alleen als de dia precies 1 afbeelding
        gebruikt (zoals de titeldia)."""
        rels_path = self._p("ppt/slides/_rels", f"slide{slide_num}.xml.rels")
        rels = open(rels_path, encoding="utf-8").read()
        m = re.search(r'Target="\.\./media/([\w.]+)"', rels)
        if not m:
            raise ValueError(f"Geen media-afbeelding gevonden voor slide{slide_num}")
        target = self._p("ppt/media", m.group(1))
        # converteer naar hetzelfde bestandsformaat als het origineel (op basis
        # van de extensie), zodat een aangeleverde .png ook goed terechtkomt
        # op een plek waar een .jpeg werd verwacht, en vice versa.
        from PIL import Image
        ext = os.path.splitext(target)[1].lower().lstrip(".")
        fmt = "JPEG" if ext in ("jpg", "jpeg") else ext.upper()
        img = Image.open(new_image_path)
        if fmt == "JPEG" and img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.save(target, format=fmt)

    def fill_rondgang_collecte(self, slide_num, doelen):
        """Vult specifiek de 3 collectedoel-waarden in na 'Eerste/Tweede/
        Derde Rondgang:' labels, zonder de rest van de dia (titel, labels,
        eventuele "(zie vlg dia)"-opmerkingen) aan te raken. Dit is
        preciezer dan replace_all_text/set_title_text, omdat de opmaak van
        deze specifieke dia's niet overal identiek is opgebouwd (de
        "(zie vlg dia)"-notitie staat niet bij elke dia op dezelfde plek)."""
        path = ("ppt/slides", f"slide{slide_num}.xml")
        xml = self._read(*path)
        runs = list(re.finditer(r"<a:t>([^<]*)</a:t>", xml))

        labels = ["Eerste Rondgang:", "Tweede Rondgang:", "Derde Rondgang:"]
        label_posities = []
        for label in labels:
            for i, m in enumerate(runs):
                if m.group(1).strip().startswith(label):
                    label_posities.append(i)
                    break

        if len(label_posities) != 3:
            raise ValueError(f"Kon niet alle 3 Rondgang-labels vinden op dia {slide_num}")

        vervangingen = {}  # run-index -> nieuwe tekst
        for idx, label_i in enumerate(label_posities):
            groep_eind = label_posities[idx + 1] if idx + 1 < len(label_posities) else len(runs)
            groep = list(range(label_i + 1, groep_eind))

            # zoek een eventuele "(...)"-opmerking binnen de groep (kan over
            # meerdere runs verdeeld zijn) -- de waarde begint pas ERNA.
            groep_tekst_per_run = [runs[i].group(1) for i in groep]
            waarde_start = 0
            open_paren = False
            for j, t in enumerate(groep_tekst_per_run):
                if "(" in t:
                    open_paren = True
                if open_paren and ")" in t:
                    waarde_start = j + 1
                    open_paren = False

            waarde_runs = groep[waarde_start:]
            if not waarde_runs:
                continue  # geen plek gevonden om de waarde te zetten; overslaan
            doel = doelen[idx] if idx < len(doelen) else ""
            vervangingen[waarde_runs[0]] = doel
            for extra in waarde_runs[1:]:
                vervangingen[extra] = ""

        # van achter naar voren vervangen zodat eerdere posities niet
        # verschuiven
        for i in sorted(vervangingen, reverse=True):
            nieuw = f"<a:t>{_esc(vervangingen[i])}</a:t>"
            m = runs[i]
            xml = xml[:m.start()] + nieuw + xml[m.end():]

        self._write(xml, *path)

    # -- nieuwe dia's met tekst (schriftlezingen) -----------------------

    def get_slide_layout(self, slide_num):
        """Welke layout (bijv. 'slideLayout4.xml') gebruikt dia N? Handig om
        nieuwe dia's dezelfde achtergrond/opmaak te geven als de dia die ze
        vervangen."""
        rels = open(self._p("ppt/slides/_rels", f"slide{slide_num}.xml.rels"), encoding="utf-8").read()
        m = re.search(r'slideLayouts/(slideLayout\d+\.xml)', rels)
        return m.group(1) if m else None

    # Tekstdia zoals de schriftlezingdia's in het sjabloon: 1 enkele
    # ctrTitle-placeholder (de layout levert de 44 pt Century Gothic en de
    # donkere achtergrond). De {paragraphs} zijn <a:p>...</a:p>-fragmenten
    # (zie bijbeltekst.regels_naar_xml).
    _TEKSTDIA_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''

    # Titeldia voor een schriftlezing: de verwijzing groot bovenaan, net
    # zoals de psalm-/lied-titeldia's in het sjabloon (60 pt Century Gothic
    # vet, wit, bovenaan de dia -- NIET half laag zoals kerkdia-txt deed).
    _LEZINGTITEL_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Tekstvak 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1056000" y="1270000"/><a:ext cx="10080000" cy="2100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="l" rtl="0"/><a:r><a:rPr lang="nl-NL" sz="6000" b="1" dirty="0"><a:solidFill><a:prstClr val="white"/></a:solidFill><a:latin typeface="Century Gothic"/></a:rPr><a:t>{titel}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''

    _TEKSTDIA_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/{layout}"/></Relationships>'''

    def _add_scaffold_slide_after(self, xml, layout, after_rid):
        """Zet een zelfgemaakte dia-XML (met alleen een layout-relatie) in
        het werkbestand en registreert 'm direct na `after_rid`."""
        n = self._next_slide_num()
        self._write(xml, "ppt/slides", f"slide{n}.xml")
        os.makedirs(self._p("ppt/slides/_rels"), exist_ok=True)
        self._write(self._TEKSTDIA_RELS_XML.format(layout=layout or "slideLayout4.xml"),
                    "ppt/slides/_rels", f"slide{n}.xml.rels")
        return self._register_slide(n, after_rid)

    def add_lezingtitel_slide_after(self, after_rid, titel, layout=None):
        """Voegt een titeldia voor één schriftlezing toe (zoals kerkdia-txt
        die maakt), direct na de dia met `after_rid`. Retourneert het nieuwe
        rId."""
        from xml.sax.saxutils import escape
        xml = self._LEZINGTITEL_XML.format(titel=escape(titel))
        return self._add_scaffold_slide_after(xml, layout, after_rid)

    def add_bijbeltekstdia_after(self, after_rid, paragraphs_xml, layout=None):
        """Voegt één tekstdia toe (1 ctrTitle-placeholder) met de gegeven
        <a:p>-fragmenten (zie bijbeltekst.regels_naar_xml). Retourneert het
        nieuwe rId."""
        xml = self._TEKSTDIA_XML.format(paragraphs=paragraphs_xml)
        return self._add_scaffold_slide_after(xml, layout, after_rid)

    def set_first_run_text_with(self, slide_num, rpr_substring, new_text, nieuwe_sz=None):
        """Vervangt de tekst van de eerste tekst-run waarvan de opmaak-attrs
        `rpr_substring` bevatten (bijv. 'sz="2800"' op de kerntekst-dia).
        Met `nieuwe_sz` (bijv. '1800') wordt óók het lettergrootte-kenmerk
        van die run aangepast (voor automatisch verkleinen van een lange
        kerntekst)."""
        from xml.sax.saxutils import escape
        path = ("ppt/slides", f"slide{slide_num}.xml")
        xml = self._read(*path)
        # zoek de eerstvolgende <a:t> ná een rPr met dit kenmerk
        pos = None
        rpr_begin = None
        for m in re.finditer(r"<a:rPr[^>]*>.*?</a:rPr>|<a:rPr[^>]*/>", xml, re.S):
            if rpr_substring in m.group(0):
                rpr_begin = m.start()
                after = xml[m.end():]
                t = re.search(r"<a:t>[^<]*</a:t>", after)
                if t:
                    pos = m.end() + t.start()
                    break
        if pos is None or rpr_begin is None:
            raise ValueError(f"geen run met {rpr_substring!r} gevonden op dia {slide_num}")
        if nieuwe_sz:
            einde_rpr = xml.index(">", rpr_begin)
            xml = (xml[:rpr_begin]
                   + re.sub(r'sz="\d+"', f'sz="{nieuwe_sz}"', xml[rpr_begin:einde_rpr], count=1)
                   + xml[einde_rpr:])
        end = xml.index("</a:t>", pos) + len("</a:t>")
        xml = xml[:pos] + f"<a:t>{escape(new_text)}</a:t>" + xml[end:]
        self._write(xml, *path)

    def strip_unused_parts(self, keep_slide_nums):
        """Optioneel: verwijdert alle slide-XML's en media die niet meer
        gebruikt worden. Handig als je maar 1-2 dia's uit een groot sjabloon
        wilt teruggeven (voorkomt een onnodig groot bestand)."""
        used_media = set()
        for n in keep_slide_nums:
            rels_path = self._p("ppt/slides/_rels", f"slide{n}.xml.rels")
            if os.path.exists(rels_path):
                used_media |= set(re.findall(r"media/([\w.]+)", open(rels_path, encoding="utf-8").read()))
        for f in os.listdir(self._p("ppt/media")):
            if f not in used_media:
                os.remove(self._p("ppt/media", f))
        for f in os.listdir(self._p("ppt/slides")):
            if f.endswith(".xml"):
                n = int(re.search(r"\d+", f).group())
                if n not in keep_slide_nums:
                    os.remove(self._p("ppt/slides", f))
        for f in os.listdir(self._p("ppt/slides/_rels")):
            n = int(re.search(r"\d+", f).group())
            if n not in keep_slide_nums:
                os.remove(self._p("ppt/slides/_rels", f))
