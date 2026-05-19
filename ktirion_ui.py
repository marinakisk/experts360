"""
ktirion_ui.py - Εκθέσεις Κτιρίων με αποθήκευση στη βάση
"""
import streamlit as st
import zipfile, re, io, tempfile
from pathlib import Path
from datetime import date
from PIL import Image

from database import (save_ekthesi_ktiriou, load_ekthesi_ktiriou,
                      search_ektheseis_ktirion, delete_ekthesi_ktiriou,
                      get_statistics_ktirion)

DAYS_GR   = ["Δευτέρα","Τρίτη","Τετάρτη","Πέμπτη","Παρασκευή","Σάββατο","Κυριακή"]
MONTHS_GR = ["Ιανουαρίου","Φεβρουαρίου","Μαρτίου","Απριλίου","Μαΐου","Ιουνίου",
              "Ιουλίου","Αυγούστου","Σεπτεμβρίου","Οκτωβρίου","Νοεμβρίου","Δεκεμβρίου"]

def fmt_long(d):
    if isinstance(d, str): return d
    return f"{DAYS_GR[d.weekday()]} {d.day} {MONTHS_GR[d.month-1]} {d.year}"

def fmt_short(d):
    if isinstance(d, str): return d
    return d.strftime("%d.%m.%Y")

def fmt_euro(v):
    return f"{v:,.2f}€".replace(",","X").replace(".",",").replace("X",".")

PAGE_W_EMU = 6200000
MAX_H_EMU  = 4500000

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
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:inline>'
        f'</w:drawing></w:r></w:p>'
    )
    if caption:
        drawing += (
            f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
            f'<w:r><w:rPr><w:i/><w:sz w:val="18"/></w:rPr>'
            f'<w:t>{caption}</w:t></w:r></w:p>'
        )
    return drawing

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
        rid_nums = [int(x) for x in re.findall(r'rId(\d+)', rels_xml) if x.isdigit()]
        rid_n  = max(rid_nums, default=30) + 1
        iid_n  = 9000

        def add_img(img_bytes):
            nonlocal rid_n, iid_n, rels_xml, ct_xml
            rid = f"rId{rid_n}"; fname = f"newimg_{rid_n}.jpg"; iid = iid_n
            rid_n += 1; iid_n += 1
            (media_dir / fname).write_bytes(img_bytes)
            rels_xml = rels_xml.replace("</Relationships>",
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{fname}"/></Relationships>')
            if 'Extension="jpg"' not in ct_xml:
                ct_xml = ct_xml.replace("</Types>", '<Default Extension="jpg" ContentType="image/jpeg"/></Types>')
            return rid, iid

        def remove_image_only_paragraphs(xml):
            def replacer(m):
                p = m.group(0)
                if '<w:drawing>' in p and not re.search(r'<w:t[ >]', p): return ''
                return p
            return re.sub(r'<w:p\b.*?</w:p>', replacer, xml, flags=re.DOTALL)

        doc_xml = remove_image_only_paragraphs(doc_xml)
        M_MAPS = "ZZZMAPSZZZ"; M_DIAP = "ZZZDIAPZZZ"; M_SYNIM = "ZZZSYNIMZZZ"

        doc_xml = re.sub(r'(<w:p\b[^>]*>(?:(?!</w:p>).)*\{\{ΓΕΝΙΚΕΣ_ΠΛΗΡΟΦΟΡΙΕΣ\}\}(?:(?!</w:p>).)*</w:p>)',
                         r'\1' + f'<w:p><w:r><w:t>{M_MAPS}</w:t></w:r></w:p>', doc_xml, flags=re.DOTALL)
        doc_xml = re.sub(r'(<w:p\b[^>]*>(?:(?!</w:p>).)*\{\{ΠΡΟΤΑΣΗ_ΑΠΟΚΑΤΑΣΤΑΣΗΣ\}\}(?:(?!</w:p>).)*</w:p>)',
                         r'\1' + f'<w:p><w:r><w:t>{M_DIAP}</w:t></w:r></w:p>', doc_xml, flags=re.DOTALL)
        doc_xml = re.sub(r'(<w:p\b[^>]*>(?:(?!</w:p>).)*ΦΩΤΟΓΡΑΦΙΚΟ ΥΛΙΚΟ(?:(?!</w:p>).)*</w:p>)',
                         r'\1' + f'<w:p><w:r><w:t>{M_SYNIM}</w:t></w:r></w:p>', doc_xml, flags=re.DOTALL)

        for ph, val in data.items():
            doc_xml = doc_xml.replace(ph, str(val))

        arim = str(data.get("{{ΑΡΙΘΜΟΣ_ΖΗΜΙΑΣ}}", ""))
        hm_ep = str(data.get("{{ΗΜ_ΕΠΙΘΕΩΡΗΣΗΣ}}", ""))
        hdr_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:p><w:pPr><w:pStyle w:val="a7"/></w:pPr>'
            '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300"'
            ' simplePos="0" relativeHeight="251658241" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>24567</wp:posOffset></wp:positionV>'
            '<wp:extent cx="1483995" cy="521970"/><wp:effectExtent l="0" t="0" r="1905" b="0"/>'
            '<wp:wrapTopAndBottom/><wp:docPr id="39" name="Logo"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="32" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1483995" cy="521970"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
            '<w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr>'
            '<w:t xml:space="preserve">Πραγματογνώμονες </w:t></w:r>'
            '<w:r><w:rPr><w:color w:val="000000"/></w:rPr>'
            f'<w:t xml:space="preserve">                                                                                                            Αρ. ζημίας:   {arim}</w:t></w:r>'
            '</w:p>'
            '<w:p><w:pPr><w:pStyle w:val="a7"/>'
            '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>'
            '<w:jc w:val="right"/></w:pPr>'
            f'<w:r><w:t>Ημ. Επιθεώρησης:  {hm_ep}</w:t></w:r></w:p>'
            '<w:p/></w:hdr>'
        )

        maps_xml = ""
        if maps_img:
            rid, iid = add_img(maps_img)
            maps_xml = img_paragraph(rid, iid, maps_img, "GoogleMaps")
        doc_xml = doc_xml.replace(f'<w:p><w:r><w:t>{M_MAPS}</w:t></w:r></w:p>', maps_xml)

        diap_xml = ""
        for i, (b, cap) in enumerate(diap_imgs):
            rid, iid = add_img(b)
            diap_xml += img_paragraph(rid, iid, b, f"Diap{i+1}", cap)
        doc_xml = doc_xml.replace(f'<w:p><w:r><w:t>{M_DIAP}</w:t></w:r></w:p>', diap_xml)

        synim_xml = ""
        for i, (b, cap) in enumerate(synim_imgs):
            rid, iid = add_img(b)
            synim_xml += img_paragraph(rid, iid, b, f"Synim{i+1}", cap)
        doc_xml = doc_xml.replace(f'<w:p><w:r><w:t>{M_SYNIM}</w:t></w:r></w:p>', synim_xml)

        doc_path.write_text(doc_xml, encoding="utf-8")
        rels_path.write_text(rels_xml, encoding="utf-8")
        hdr_path.write_text(hdr_xml, encoding="utf-8")
        ct_path.write_text(ct_xml, encoding="utf-8")

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for p in sorted(tmp.rglob("*")):
                if p.is_file():
                    zout.write(p, p.relative_to(tmp).as_posix())
        return out.getvalue()


def show_ktirion_tab(page="📝 Νέα Έκθεση"):

    if page == "🔍 Αναζήτηση":
        st.subheader("🔍 Αναζήτηση Εκθέσεων Κτιρίων")
        qk = st.text_input("Αναζήτηση (αρ.ζημίας, παθών, τοποθεσία...)")
        sk = st.selectbox("Status", ["","draft","final","archived"],
                          format_func=lambda x: {"":"Όλα","draft":"Προσχέδιο","final":"Τελική","archived":"Αρχείο"}.get(x,x),
                          key="ktir_status_filter")
        results_k = search_ektheseis_ktirion(query=qk, status=sk)
        st.write(f"**{len(results_k)} εκθέσεις κτιρίων**")
        for r in results_k:
            sc = {"draft":"🟡","final":"🟢","archived":"⚫"}.get(r.get('status',''),'⚪')
            with st.expander(f"{sc} #{r['id']} | {r.get('arithmos_zimias','')} | {r.get('pathon','')} | {r.get('topothesia','')}"):
                ic1,ic2,ic3 = st.columns(3)
                ic1.caption(f"**Ημ. Επιθεώρησης:** {r.get('hm_epitheorisis','—')}")
                ic2.caption(f"**Απαίτηση:** {r.get('total_apaitisi',0):,.2f} €")
                ic3.caption(f"**Εκτίμηση:** {r.get('total_ektimisi',0):,.2f} €")
                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button("📂 Φόρτωση", key=f"kload_{r['id']}", use_container_width=True):
                        ek = load_ekthesi_ktiriou(r['id'])
                        if ek:
                            st.session_state['current_ktirion_id'] = r['id']
                            st.session_state['ktir_loaded'] = ek
                            st.rerun()
                with bc2:
                    if st.button("🗑️ Διαγραφή", key=f"kdel_{r['id']}", use_container_width=True):
                        ok, _ = delete_ekthesi_ktiriou(r['id'])
                        if ok: st.rerun()
        return

    if page == "📊 Στατιστικά":
        st.subheader("📊 Στατιστικά Κτιρίων")
        sk = get_statistics_ktirion()
        c1,c2,c3 = st.columns(3)
        c1.metric("Σύνολο εκθέσεων", sk.get('total',0))
        c2.metric("Τελευταίες 30 μέρες", sk.get('last_30_days',0))
        c3.metric("Συνολική εκτίμηση", f"{sk.get('total_ektimisi',0):,.0f} €")
        return

    # ── ΝΕΑ ΕΚΘΕΣΗ ───────────────────────────────────────────
    loaded = st.session_state.pop('ktir_loaded', None)
    eid = st.session_state.get('current_ktirion_id')

    if eid:
        st.info(f"📂 Επεξεργασία έκθεσης #{eid}")
        if st.button("➕ Νέα Έκθεση", key="ktir_new"):
            for k in ['current_ktirion_id','damage_rows','bullets','ek_override','diap_captions','synim_captions']:
                st.session_state.pop(k, None)
            st.rerun()

    template_bytes = None
    tp = Path("template_ekthesis.docx")
    if tp.exists():
        template_bytes = tp.read_bytes()
    else:
        up_tpl = st.file_uploader("📂 Ανέβασε το template_ekthesis.docx", type=["docx"])
        if up_tpl:
            template_bytes = up_tpl.read()
    if not template_bytes:
        st.warning("⚠️ Βάλε το template_ekthesis.docx στον ίδιο φάκελο με το app.")

    if 'damage_rows' not in st.session_state:
        st.session_state.damage_rows = [{"desc": "", "apaitisi": 0.0}]
    if 'bullets' not in st.session_state:
        st.session_state.bullets = [""]
    if 'ek_override' not in st.session_state:
        st.session_state.ek_override = {}
    if 'diap_captions' not in st.session_state:
        st.session_state.diap_captions = {}
    if 'synim_captions' not in st.session_state:
        st.session_state.synim_captions = {}

    if loaded:
        arithmos_def      = loaded.get('arithmos_zimias','')
        pathon_def        = loaded.get('pathon','')
        up_opsin_def      = loaded.get('up_opsin','')
        asfalismen_def    = loaded.get('asfalismenos','')
        asfalister_def    = loaded.get('asfalisterio','')
        topothesia_def    = loaded.get('topothesia','')
        thema_def         = loaded.get('thema','')
        eisagogi_def      = loaded.get('eisagogi','')
        genikes_def       = loaded.get('genikes','')
        eisagogi_diap_def = loaded.get('eisagogi_diap','')
        protasi_def       = loaded.get('protasi','')
        prosopiki_def     = loaded.get('prosopiki','')
        apaitisi_par_def  = loaded.get('apaitisi_par','')
        grammes = loaded.get('grammes', [])
        if grammes:
            st.session_state.damage_rows = [{"desc": g.get('desc',''), "apaitisi": float(g.get('apaitisi',0))} for g in grammes]
            st.session_state.ek_override = {i: float(g.get('ektimisi',0)) for i, g in enumerate(grammes)}
        bullets_raw = loaded.get('bullets','')
        if bullets_raw:
            st.session_state.bullets = bullets_raw.split('\n')
    else:
        arithmos_def = pathon_def = up_opsin_def = asfalismen_def = asfalister_def = topothesia_def = thema_def = ''
        eisagogi_def = "επικοινωνήσαμε με την παθούσα, ώστε να ενημερωθούμε για το ακριβές σημείο ζημιάς"
        genikes_def = ''
        eisagogi_diap_def = "Κατά την επιθεώρησή μας στο σημείο ζημίας, διαπιστώσαμε τα κάτωθι:"
        protasi_def = "Για την αποκατάσταση των προκληθέντων ζημιών θα πρέπει να αντικατασταθεί "
        prosopiki_def = ''
        apaitisi_par_def = "Για την εξεταζόμενη ζημία, ο παθών μας απέστειλε την οικονομική προσφορά αντί απαιτήσεως."

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("## 🔢 Στοιχεία Ζημίας")
        arithmos   = st.text_input("Αριθμός Ζημίας *", value=arithmos_def, placeholder="π.χ. 39150000")
        hm_epith   = st.date_input("Ημ. Επιθεώρησης *", value=date.today())
        hm_synt    = st.date_input("Ημ. Σύνταξης Έκθεσης", value=date.today())
        hm_anath   = st.date_input("Ημ. Ανάθεσης Εντολής", value=date.today())
        hm_zim     = st.date_input("Χρόνος Ζημίας", value=date.today())
    with col2:
        st.markdown("## 👤 Στοιχεία Παθόντος")
        pathon     = st.text_input("Παθών *", value=pathon_def, placeholder="π.χ. ΠΑΠΑΔΟΠΟΥΛΟΥ ΜΑΡΙΑ")
        up_opsin   = st.text_input("Υπ' όψιν", value=up_opsin_def)
        asfalismen = st.text_input("Ασφαλισμένος / Αρ. Πινακίδας", value=asfalismen_def)
        asfalister = st.text_input("Ασφαλιστήριο Συμβόλαιο", value=asfalister_def)
        topothesia = st.text_input("Τοποθεσία Ζημίας *", value=topothesia_def)

    st.markdown("## ✍️ Περιγραφές & Φωτογραφίες")
    thema = st.text_input("Θέμα Ζημίας", value=thema_def)

    col3, col4 = st.columns(2)
    with col3:
        eisagogi = st.text_area("Εισαγωγική Παράγραφος", value=eisagogi_def, height=110)
        st.markdown("**Γενικές Πληροφορίες**")
        genikes = st.text_area("genikes", label_visibility="collapsed", value=genikes_def, height=90)
        st.markdown("📍 **Google Maps φωτό**")
        maps_file = st.file_uploader("maps", type=["jpg","jpeg","png"], label_visibility="collapsed", key="maps_up")
        if maps_file:
            st.image(maps_file, use_container_width=True)
    with col4:
        eisagogi_diap = st.text_input("Εισαγωγή Διαπιστώσεων", value=eisagogi_diap_def)
        st.markdown("**Διαπιστώσεις (bullets)**")
        for bi in range(len(st.session_state.bullets)):
            b1, b2 = st.columns([10, 1])
            with b1:
                st.session_state.bullets[bi] = st.text_input(
                    f"b{bi}", st.session_state.bullets[bi],
                    label_visibility="collapsed", key=f"bul_{bi}")
            with b2:
                if st.button("✕", key=f"xb_{bi}") and len(st.session_state.bullets) > 1:
                    st.session_state.bullets.pop(bi)
                    st.rerun()
        if st.button("➕ Bullet"):
            st.session_state.bullets.append("")
            st.rerun()
        protasi   = st.text_area("Πρόταση Αποκατάστασης", value=protasi_def, height=85)
        prosopiki = st.text_area("Προσωπική Άποψη", value=prosopiki_def, height=60)

    st.markdown("### 📸 Φωτό Διαπιστώσεων")
    diap_files = st.file_uploader("diap", type=["jpg","jpeg","png"],
                                  accept_multiple_files=True, key="diap_up", label_visibility="collapsed")
    if diap_files:
        if len(diap_files) > 3:
            diap_files = diap_files[:3]
        for i, f in enumerate(diap_files):
            c1, c2 = st.columns([2, 3])
            with c1: st.image(f, caption=f"Διαπίστωση {i+1}", use_container_width=True)
            with c2:
                st.session_state.diap_captions[i] = st.text_input(
                    f"Λεζάντα {i+1}", value=st.session_state.diap_captions.get(i,""), key=f"dcap_{i}")

    apaitisi_par = st.text_area("Παράγραφος Απαίτησης", value=apaitisi_par_def, height=100)

    st.markdown("## 💶 Πίνακας Κόστους")

    def add_row_k(): st.session_state.damage_rows.append({"desc": "", "apaitisi": 0.0})
    def del_row_k(i):
        if len(st.session_state.damage_rows) > 1:
            st.session_state.damage_rows.pop(i)
            st.session_state.ek_override.pop(i, None)

    hc = st.columns([5, 2, 2, 1])
    hc[0].markdown("**Περιγραφή**"); hc[1].markdown("**Απαίτηση (€)**"); hc[2].markdown("**Εκτίμηση (€)**")

    total_ap = 0.0; total_ek = 0.0; rows_data = []
    for i, row in enumerate(st.session_state.damage_rows):
        rc = st.columns([5, 2, 2, 1])
        with rc[0]:
            desc = st.text_input(f"d{i}", value=row["desc"], label_visibility="collapsed", key=f"desc_{i}")
        with rc[1]:
            ap = st.number_input("a", value=float(row["apaitisi"]), min_value=0.0, step=10.0,
                                 format="%.2f", label_visibility="collapsed", key=f"ap_{i}")
        ek_def = st.session_state.ek_override.get(i, ap)
        with rc[2]:
            ek = st.number_input("e", value=float(ek_def), min_value=0.0, step=10.0,
                                 format="%.2f", label_visibility="collapsed", key=f"ek_{i}")
            st.session_state.ek_override[i] = ek
        with rc[3]:
            if st.button("🗑️", key=f"del_{i}"): del_row_k(i); st.rerun()
        st.session_state.damage_rows[i] = {"desc": desc, "apaitisi": ap}
        total_ap += ap; total_ek += ek
        rows_data.append({"desc": desc, "apaitisi": ap, "ektimisi": ek})

    st.button("➕ Προσθήκη γραμμής", on_click=add_row_k)

    fpa_ap = total_ap * 0.24; fpa_ek = total_ek * 0.24
    tel_ap = total_ap + fpa_ap; tel_ek = total_ek + fpa_ek

    st.markdown(f"""<div style="background:#e8f5e9;border-radius:8px;padding:1rem;margin-top:1rem;">
<table style="width:100%;border-collapse:collapse;">
<tr><td style="width:55%"></td><th style="text-align:right;padding:4px 16px;">Απαίτηση</th><th style="text-align:right;padding:4px 16px;">Εκτίμηση</th></tr>
<tr><td>Μερικό σύνολο</td><td style="text-align:right;padding:4px 16px;">{fmt_euro(total_ap)}</td><td style="text-align:right;padding:4px 16px;">{fmt_euro(total_ek)}</td></tr>
<tr><td>ΦΠΑ 24%</td><td style="text-align:right;padding:4px 16px;">{fmt_euro(fpa_ap)}</td><td style="text-align:right;padding:4px 16px;">{fmt_euro(fpa_ek)}</td></tr>
<tr style="font-weight:bold;border-top:2px solid #388e3c;"><td>ΣΥΝΟΛΟ</td><td style="text-align:right;padding:4px 16px;">{fmt_euro(tel_ap)}</td><td style="text-align:right;padding:4px 16px;">{fmt_euro(tel_ek)}</td></tr>
</table></div>""", unsafe_allow_html=True)

    st.markdown("## 🖼️ Φωτό Συννημένων")
    synim_files = st.file_uploader("synim", type=["jpg","jpeg","png"],
                                   accept_multiple_files=True, key="synim_up", label_visibility="collapsed")
    if synim_files:
        for i, f in enumerate(synim_files):
            c1, c2 = st.columns([2, 3])
            with c1: st.image(f, caption=f"Φωτό {i+1}", use_container_width=True)
            with c2:
                st.session_state.synim_captions[i] = st.text_input(
                    f"Λεζάντα {i+1}", value=st.session_state.synim_captions.get(i,""), key=f"scap_{i}")

    # ΑΠΟΘΗΚΕΥΣΗ
    st.markdown("---")
    st.subheader("💾 Αποθήκευση")
    sc1, sc2, sc3 = st.columns([2,2,2])
    with sc1:
        save_user_k = st.text_input("Όνομα χρήστη", key="ktir_save_user", placeholder="π.χ. Μαρινάκης")
    with sc2:
        save_status_k = st.selectbox("Status", ["draft","final","archived"],
                                     format_func=lambda x: {"draft":"Προσχέδιο","final":"Τελική","archived":"Αρχείο"}[x],
                                     key="ktir_save_status")
    with sc3:
        st.write(""); st.write("")
        btn_lbl = f"💾 Ενημέρωση #{eid}" if eid else "💾 Αποθήκευση νέας"
        if st.button(btn_lbl, type="primary", use_container_width=True, key="ktir_save_btn"):
            save_data_k = {
                "arithmos_zimias": arithmos, "hm_epitheorisis": str(hm_epith),
                "hm_syntaxis": str(hm_synt), "hm_anathesis": str(hm_anath),
                "hm_zimias": str(hm_zim), "pathon": pathon, "up_opsin": up_opsin,
                "asfalismenos": asfalismen, "asfalisterio": asfalister,
                "topothesia": topothesia, "thema": thema, "eisagogi": eisagogi,
                "genikes": genikes, "eisagogi_diap": eisagogi_diap,
                "bullets": "\n".join(st.session_state.bullets),
                "protasi": protasi, "prosopiki": prosopiki, "apaitisi_par": apaitisi_par,
                "total_apaitisi": total_ap, "total_ektimisi": total_ek,
                "asfalistiki": st.session_state.get('asfalistiki',''),
                "onomateponymo": save_user_k, "status": save_status_k,
            }
            new_id_k, err_k = save_ekthesi_ktiriou(
                data=save_data_k, grammes=rows_data, user_name=save_user_k, ekthesi_id=eid)
            if err_k:
                st.error(f"❌ {err_k}")
            else:
                st.session_state['current_ktirion_id'] = new_id_k
                st.success(f"✅ Αποθηκεύτηκε #{new_id_k}")

    # ΠΑΡΑΓΩΓΗ DOCX
    st.markdown("---")
    st.subheader("📄 Παραγωγή Εκθέσεως")
    fn_default = (f"ekthesi_{arithmos}_{fmt_short(hm_synt).replace('.','')}.docx" if arithmos else "nea_ekthesi.docx")
    out_name = st.text_input("Όνομα αρχείου εξόδου", value=fn_default)

    if st.button("🖨️ Δημιουργία Εκθέσεως", type="primary", use_container_width=True):
        if not template_bytes:
            st.error("❌ Δεν βρέθηκε το template_ekthesis.docx")
        else:
            errs = []
            if not arithmos: errs.append("Αριθμός Ζημίας")
            if not pathon:   errs.append("Παθών")
            if not topothesia: errs.append("Τοποθεσία")
            if all(r["desc"] == "" for r in rows_data): errs.append("Τουλάχιστον μία γραμμή ζημίας")
            if errs:
                st.error(f"⚠️ Συμπλήρωσε: {', '.join(errs)}")
            else:
                MAX = 10; tbl = {}
                for idx in range(MAX):
                    n = idx + 1
                    if idx < len(rows_data):
                        r = rows_data[idx]
                        tbl[f"{{{{ΕΙΔΟΣ_{n}}}}}"]    = r["desc"]
                        tbl[f"{{{{ΑΠΑΙΤΗΣΗ_{n}}}}}"] = fmt_euro(r["apaitisi"])
                        tbl[f"{{{{ΕΚΤΙΜΗΣΗ_{n}}}}}"] = fmt_euro(r["ektimisi"])
                    else:
                        tbl[f"{{{{ΕΙΔΟΣ_{n}}}}}"] = tbl[f"{{{{ΑΠΑΙΤΗΣΗ_{n}}}}}"] = tbl[f"{{{{ΕΚΤΙΜΗΣΗ_{n}}}}}"] = ""

                bullets_text = "\n".join(b for b in st.session_state.bullets if b.strip())
                data = {
                    "{{ΑΡΙΘΜΟΣ_ΖΗΜΙΑΣ}}": arithmos, "{{ΗΜ_ΕΠΙΘΕΩΡΗΣΗΣ}}": fmt_short(hm_epith),
                    "{{ΗΜΕΡΟΜΗΝΙΑ_ΣΥΝΤΑΞΗΣ}}": fmt_long(hm_synt), "{{ΘΕΜΑ_ΖΗΜΙΑΣ}}": thema,
                    "{{ΥΠ_ΟΨΙΝ}}": up_opsin, "{{ΑΣΦΑΛΙΣΜΕΝΟΣ}}": asfalismen,
                    "{{ΑΣΦΑΛΙΣΤΗΡΙΟ}}": asfalister, "{{ΠΑΘΩΝ}}": pathon,
                    "{{ΤΟΠΟΘΕΣΙΑ_ΖΗΜΙΑΣ}}": topothesia, "{{ΧΡΟΝΟΣ_ΖΗΜΙΑΣ}}": fmt_short(hm_zim),
                    "{{ΗΜ_ΑΝΑΘΕΣΗΣ}}": fmt_short(hm_anath),
                    "{{ΑΠΑΙΤΗΣΗ_ΣΥΝΟΛΟ}}": fmt_euro(total_ap), "{{ΕΚΤΙΜΗΣΗ_ΣΥΝΟΛΟ}}": fmt_euro(total_ek),
                    "{{ΕΙΣΑΓΩΓΗ_ΠΕΡΙΓΡΑΦΗ}}": eisagogi, "{{ΓΕΝΙΚΕΣ_ΠΛΗΡΟΦΟΡΙΕΣ}}": genikes,
                    "{{ΕΙΣΑΓΩΓΗ_ΔΙΑΠΙΣΤΩΣΕΩΝ}}": eisagogi_diap, "{{ΔΙΑΠΙΣΤΩΣΕΙΣ_BULLETS}}": bullets_text,
                    "{{ΠΡΟΤΑΣΗ_ΑΠΟΚΑΤΑΣΤΑΣΗΣ}}": protasi, "{{ΠΡΟΣΩΠΙΚΗ_ΑΠΟΨΗ}}": prosopiki,
                    "{{ΑΠΑΙΤΗΣΗ_ΠΕΡΙΓΡΑΦΗ}}": apaitisi_par, **tbl,
                    "{{ΜΕΡΙΚΟ_ΣΥΝΟΛΟ_ΑΠΑΙΤΗΣΗ}}": fmt_euro(total_ap), "{{ΜΕΡΙΚΟ_ΣΥΝΟΛΟ_ΕΚΤΙΜΗΣΗ}}": fmt_euro(total_ek),
                    "{{ΦΠΑ_ΑΠΑΙΤΗΣΗ}}": fmt_euro(fpa_ap), "{{ΦΠΑ_ΕΚΤΙΜΗΣΗ}}": fmt_euro(fpa_ek),
                    "{{ΣΥΝΟΛΟ_ΑΠΑΙΤΗΣΗ}}": fmt_euro(tel_ap), "{{ΣΥΝΟΛΟ_ΕΚΤΙΜΗΣΗ}}": fmt_euro(tel_ek),
                }
                maps_b  = to_jpeg(maps_file) if maps_file else None
                diap_b  = [(to_jpeg(f), st.session_state.diap_captions.get(i,"")) for i, f in enumerate(diap_files or [])]
                synim_b = [(to_jpeg(f), st.session_state.synim_captions.get(i,"")) for i, f in enumerate(synim_files or [])]

                with st.spinner("Παράγεται η έκθεση..."):
                    try:
                        result = fill_template(template_bytes, data, maps_b, diap_b, synim_b)
                        st.success("✅ Η έκθεση δημιουργήθηκε!")
                        st.download_button(
                            label=f"⬇️ Κατέβασε: {out_name}", data=result,
                            file_name=out_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary", use_container_width=True)
                    except Exception as e:
                        st.error(f"❌ Σφάλμα: {e}")
                        import traceback; st.code(traceback.format_exc())