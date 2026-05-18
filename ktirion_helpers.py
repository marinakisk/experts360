"""
ktirion_helpers.py - Βοηθητικές συναρτήσεις για εκθέσεις κτιρίων
"""
import zipfile, re, io, tempfile
from pathlib import Path
from datetime import date
from PIL import Image

"""
=============================================================
 ΠΡΑΓΜΑΤΟΓΝΩΜΟΣΥΝΗ - Φόρμα Παραγωγής Εκθέσεων
 Experts360
=============================================================

ΧΡΗΣΗ:
1. pip install streamlit pillow
2. streamlit run app.py
3. Βάλε το template_ekthesis.docx στον ίδιο φάκελο
"""


# =============================================================
# CONFIG
# =============================================================





# =============================================================
# HELPERS – Ημερομηνίες & Ποσά
# =============================================================

DAYS_GR   = ["Δευτέρα","Τρίτη","Τετάρτη","Πέμπτη","Παρασκευή","Σάββατο","Κυριακή"]
MONTHS_GR = ["Ιανουαρίου","Φεβρουαρίου","Μαρτίου","Απριλίου","Μαΐου","Ιουνίου",
              "Ιουλίου","Αυγούστου","Σεπτεμβρίου","Οκτωβρίου","Νοεμβρίου","Δεκεμβρίου"]

def fmt_long(d):
    return f"{DAYS_GR[d.weekday()]} {d.day} {MONTHS_GR[d.month-1]} {d.year}"

def fmt_short(d):
    return d.strftime("%d.%m.%Y")

def fmt_euro(v):
    return f"{v:,.2f}€".replace(",","X").replace(".",",").replace("X",".")

# =============================================================
# HELPERS – Εικόνες
# =============================================================

PAGE_W_EMU = 6200000   # ~17.2cm, χωράει στη σελίδα
MAX_H_EMU  = 4500000   # ~12.5cm max ύψος

def to_jpeg(f):
    img = Image.open(f)
    if img.mode in ("RGBA","P"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()

def emu_size(img_bytes):
    img = Image.open(io.BytesIO(img_bytes))
    w, h = img.size
    cx = PAGE_W_EMU
    cy = int(cx * h / w)
    if cy > MAX_H_EMU:
        cy = MAX_H_EMU
        cx = int(cy * w / h)
    return cx, cy

def img_paragraph(rid, iid, img_bytes, name="Photo", caption=""):
    cx, cy = emu_size(img_bytes)
    drawing = (
        f'<w:p><w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{iid}" name="{name}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{iid}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{rid}" cstate="print"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing></w:r></w:p>'
    )
    if caption and caption.strip():
        cap_xml = (
            f'<w:p>'
            f'<w:pPr><w:jc w:val="center"/>'
            f'<w:spacing w:before="40" w:after="120"/></w:pPr>'
            f'<w:r>'
            f'<w:rPr><w:i/><w:iCs/><w:sz w:val="18"/><w:szCs w:val="18"/>'
            f'<w:color w:val="595959"/></w:rPr>'
            f'<w:t xml:space="preserve">{caption}</w:t>'
            f'</w:r>'
            f'</w:p>'
        )
        return drawing + cap_xml
    return drawing

def merge_split_placeholders(xml):
    """
    Το Word σπάει {{PLACEHOLDER}} σε χωριστά runs.
    Γρήγορη λύση: αφαίρεσε τα </w:r><w:r...> ανάμεσα σε τμήματα του placeholder.
    """
    # Βρες και ένωσε: </w:t></w:r><w:r...><w:t...> που μαζί σχηματίζουν {{...}}
    # Απλή προσέγγιση: concatenate όλα τα w:t μέσα σε κάθε w:p, βρες placeholders,
    # και αντικατάστησε το split pattern με ένα ενιαίο run.

    def fix_para(m):
        para = m.group(0)
        if '{{' not in para:
            return para

        # Μάζεψε όλα τα consecutive <w:t> values με τα surrounding runs
        # Χτίσε ένα "flat text" map
        parts = re.split(r'(<w:r\b[^>]*>|</w:r>|<w:t[^>]*>|</w:t>)', para)

        # Πιο απλή προσέγγιση: βρες sequences of runs που span ένα placeholder
        # Ένωσε απλά το text content αγνοώντας run boundaries μέσα σε {{...}}
        # Κάνε replace στο concatenated text και rebuild

        # Extract run blocks
        run_blocks = re.findall(r'<w:r\b.*?</w:r>', para, re.DOTALL)
        if len(run_blocks) < 2:
            return para

        # Βρες groups που μαζί κάνουν ένα placeholder
        i = 0
        new_para = para
        while i < len(run_blocks):
            t = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', run_blocks[i]))
            if '{{' in t and '}}' not in t:
                # Ξεκινάει placeholder — μάζεψε επόμενα runs
                j = i + 1
                combined = t
                group = [run_blocks[i]]
                while j < len(run_blocks) and '}}' not in combined:
                    nt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', run_blocks[j]))
                    combined += nt
                    group.append(run_blocks[j])
                    j += 1
                if '}}' in combined:
                    # Συγχώνευση
                    rpr = re.search(r'<w:rPr>.*?</w:rPr>', group[0], re.DOTALL)
                    rpr_xml = rpr.group(0) if rpr else ''
                    merged = f'<w:r>{rpr_xml}<w:t xml:space="preserve">{combined}</w:t></w:r>'
                    old = ''.join(group)
                    new_para = new_para.replace(old, merged, 1)
                    run_blocks = re.findall(r'<w:r\b.*?</w:r>', new_para, re.DOTALL)
            i += 1
        return new_para

    return re.sub(r'<w:p\b.*?</w:p>', fix_para, xml, flags=re.DOTALL)



def fill_template(template_bytes, data, maps_img, diap_imgs, synim_imgs):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        with zipfile.ZipFile(io.BytesIO(template_bytes)) as z:
            z.extractall(tmp)

        doc_path  = tmp / "word" / "document.xml"
        rels_path = tmp / "word" / "_rels" / "document.xml.rels"
        hdr_path  = tmp / "word" / "header1.xml"
        ct_path   = tmp / "[Content_Types].xml"
        media_dir = tmp / "word" / "media"
        media_dir.mkdir(exist_ok=True)

        doc_xml  = doc_path.read_text(encoding="utf-8")
        rels_xml = rels_path.read_text(encoding="utf-8")
        hdr_xml  = hdr_path.read_text(encoding="utf-8")
        ct_xml   = ct_path.read_text(encoding="utf-8")

        # --- Μετρητές ---
        rid_nums = [int(x) for x in re.findall(r'rId(\d+)', rels_xml) if x.isdigit()]
        rid_n  = max(rid_nums, default=30) + 1
        iid_n  = 9000

        def add_img(img_bytes):
            nonlocal rid_n, iid_n, rels_xml, ct_xml
            rid   = f"rId{rid_n}"
            fname = f"newimg_{rid_n}.jpg"
            iid   = iid_n
            rid_n  += 1
            iid_n  += 1
            (media_dir / fname).write_bytes(img_bytes)
            rels_xml = rels_xml.replace("</Relationships>",
                f'<Relationship Id="{rid}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="media/{fname}"/></Relationships>')
            if 'Extension="jpg"' not in ct_xml:
                ct_xml = ct_xml.replace("</Types>",
                    '<Default Extension="jpg" ContentType="image/jpeg"/></Types>')
            return rid, iid

        # ── ΒΗΜΑ 1: Αφαίρεση original εικόνων του template ──────────
        # Τα paragraphs που έχουν ΜΟΝΟ drawing (χωρίς text) αφαιρούνται
        # Συγκεκριμένα: para 47,48 (μετά ΓΕΝΙΚΕΣ), para 57 (μετά BULLETS), para 112 (ΣΥΝΝΗΜΕΝΑ)
        # Τα αναγνωρίζουμε: <w:p> που περιέχει <w:drawing> αλλά ΟΧΙ <w:t>
        def remove_image_only_paragraphs(xml):
            """Αφαιρεί paragraphs που περιέχουν μόνο εικόνες (χωρίς κείμενο) από το template."""
            def replacer(m):
                p = m.group(0)
                if '<w:drawing>' in p and not re.search(r'<w:t[ >]', p):
                    return ''  # Αφαίρεση
                return p
            return re.sub(r'<w:p\b.*?</w:p>', replacer, xml, flags=re.DOTALL)

        doc_xml = remove_image_only_paragraphs(doc_xml)

        # ── ΒΗΜΑ 2: Markers για εισαγωγή νέων εικόνων ───────────────
        M_MAPS  = "ZZZMAPSZZZ"
        M_DIAP  = "ZZZDIAPZZZ"
        M_SYNIM = "ZZZSYNIMZZZ"

        # Maps: μετά το paragraph ΓΕΝΙΚΕΣ_ΠΛΗΡΟΦΟΡΙΕΣ
        doc_xml = re.sub(
            r'(<w:p\b[^>]*>(?:(?!</w:p>).)*\{\{ΓΕΝΙΚΕΣ_ΠΛΗΡΟΦΟΡΙΕΣ\}\}(?:(?!</w:p>).)*</w:p>)',
            r'\1' + f'<w:p><w:r><w:t>{M_MAPS}</w:t></w:r></w:p>',
            doc_xml, flags=re.DOTALL
        )
        # Diap: μετά το paragraph ΠΡΟΤΑΣΗ_ΑΠΟΚΑΤΑΣΤΑΣΗΣ (όχι μέσα στα bullets)
        doc_xml = re.sub(
            r'(<w:p\b[^>]*>(?:(?!</w:p>).)*\{\{ΠΡΟΤΑΣΗ_ΑΠΟΚΑΤΑΣΤΑΣΗΣ\}\}(?:(?!</w:p>).)*</w:p>)',
            r'\1' + f'<w:p><w:r><w:t>{M_DIAP}</w:t></w:r></w:p>',
            doc_xml, flags=re.DOTALL
        )
        # Synim: μετά ΦΩΤΟΓΡΑΦΙΚΟ ΥΛΙΚΟ
        doc_xml = re.sub(
            r'(<w:p\b[^>]*>(?:(?!</w:p>).)*ΦΩΤΟΓΡΑΦΙΚΟ ΥΛΙΚΟ(?:(?!</w:p>).)*</w:p>)',
            r'\1' + f'<w:p><w:r><w:t>{M_SYNIM}</w:t></w:r></w:p>',
            doc_xml, flags=re.DOTALL
        )

        # ── ΒΗΜΑ 3: Text replace (document + header) ─────────────────
        for ph, val in data.items():
            doc_xml = doc_xml.replace(ph, str(val))

        # Header: χτίζουμε καινούριο καθαρό XML με τις τιμές απευθείας
        # Κρατάμε μόνο το λογότυπο (anchor με rId1) και γράφουμε τα πεδία
        arim = str(data.get("{{ΑΡΙΘΜΟΣ_ΖΗΜΙΑΣ}}", ""))
        hm_ep = str(data.get("{{ΗΜ_ΕΠΙΘΕΩΡΗΣΗΣ}}", ""))

        hdr_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">'

            # Γραμμή 1: Λογότυπο (floating) + "Πραγματογνώμονες" + "Αρ. ζημίας: ΤΙΜΗ"
            '<w:p>'
            '<w:pPr><w:pStyle w:val="a7"/></w:pPr>'
            # Λογότυπο ως floating anchor
            '<w:r>'
            '<w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            '<w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300"'
            ' simplePos="0" relativeHeight="251658241" behindDoc="1" locked="0"'
            ' layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>24567</wp:posOffset></wp:positionV>'
            '<wp:extent cx="1483995" cy="521970"/>'
            '<wp:effectExtent l="0" t="0" r="1905" b="0"/>'
            '<wp:wrapTopAndBottom/>'
            '<wp:docPr id="39" name="Logo"/>'
            '<wp:cNvGraphicFramePr/>'
            '<a:graphic>'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="32" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill>'
            '<a:blip r:embed="rId1"/>'
            '<a:stretch><a:fillRect/></a:stretch>'
            '</pic:blipFill>'
            '<pic:spPr>'
            '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1483995" cy="521970"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '</pic:spPr>'
            '</pic:pic>'
            '</a:graphicData>'
            '</a:graphic>'
            '</wp:anchor>'
            '</w:drawing>'
            '</w:r>'
            # Κείμενο "Πραγματογνώμονες" + tabs + "Αρ. ζημίας: ΤΙΜΗ"
            '<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/>'
            '<w:color w:val="000000"/></w:rPr>'
            '<w:t xml:space="preserve">Πραγματογνώμονες </w:t></w:r>'
            '<w:r><w:rPr><w:color w:val="000000"/></w:rPr>'
            f'<w:t xml:space="preserve">                                                                                                            Αρ. ζημίας:   {arim}</w:t></w:r>'
            '</w:p>'

            # Γραμμή 2: "Ημ. Επιθεώρησης: ΤΙΜΗ" (δεξιά)
            '<w:p>'
            '<w:pPr><w:pStyle w:val="a7"/>'
            '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>'
            '<w:jc w:val="right"/></w:pPr>'
            '<w:r><w:rPr><w:color w:val="000000"/></w:rPr>'
            f'<w:t>Ημ. Επιθεώρησης:  {hm_ep}</w:t></w:r>'
            '</w:p>'

            '<w:p/>'
            '</w:hdr>'
        )

        # ── ΒΗΜΑ 4: Αντικατάσταση markers με εικόνες ────────────────

        # Maps (0 ή 1)
        if maps_img:
            rid, iid = add_img(maps_img)
            maps_xml = img_paragraph(rid, iid, maps_img, "GoogleMaps")
        else:
            maps_xml = ""
        doc_xml = doc_xml.replace(f'<w:p><w:r><w:t>{M_MAPS}</w:t></w:r></w:p>', maps_xml)

        # Diap (1-3, κάτω-κάτω)
        diap_xml = ""
        for i, (b, cap) in enumerate(diap_imgs):
            rid, iid = add_img(b)
            diap_xml += img_paragraph(rid, iid, b, f"Diap{i+1}", cap)
        doc_xml = doc_xml.replace(f'<w:p><w:r><w:t>{M_DIAP}</w:t></w:r></w:p>', diap_xml)

        # Synim (όσες θέλεις)
        synim_xml = ""
        for i, (b, cap) in enumerate(synim_imgs):
            rid, iid = add_img(b)
            synim_xml += img_paragraph(rid, iid, b, f"Synim{i+1}", cap)
        doc_xml = doc_xml.replace(f'<w:p><w:r><w:t>{M_SYNIM}</w:t></w:r></w:p>', synim_xml)

        # ── ΒΗΜΑ 5: Αποθήκευση ───────────────────────────────────────
        doc_path.write_text(doc_xml, encoding="utf-8")
        rels_path.write_text(rels_xml, encoding="utf-8")
        hdr_path.write_text(hdr_xml, encoding="utf-8")
        ct_path.write_text(ct_xml, encoding="utf-8")

        # Επαλήθευση ότι το header γράφτηκε σωστά
        check = hdr_path.read_text(encoding="utf-8")
        for ph in ["{{ΑΡΙΘΜΟΣ_ΖΗΜΙΑΣ}}", "{{ΗΜ_ΕΠΙΘΕΩΡΗΣΗΣ}}"]:
            if ph in check:
                # Κάτι πήγε στραβά — κάνε replace ξανά απευθείας στο αρχείο
                check = check.replace(ph, str(data.get(ph, "")))
        hdr_path.write_text(check, encoding="utf-8")

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for p in sorted(tmp.rglob("*")):
                if p.is_file():
                    arcname = p.relative_to(tmp).as_posix()
                    zout.write(p, arcname)
        return out.getvalue()


