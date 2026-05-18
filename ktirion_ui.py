import streamlit as st
from ktirion_helpers import (fmt_long, fmt_short, fmt_euro, to_jpeg,
                              img_paragraph, merge_split_placeholders, fill_template)
from database import (save_ekthesi_ktiriou, load_ekthesi_ktiriou,
                       search_ektheseis_ktirion, delete_ekthesi_ktiriou,
                       get_statistics_ktirion)
from pathlib import Path
from datetime import date
import io


def show_ktirion_tab():
    """Εμφανίζει το tab εκθέσεων κτιρίων."""

    # =============================================================
    # SESSION STATE
    # =============================================================
    
    if "damage_rows" not in st.session_state:
        st.session_state.damage_rows = [{"desc": "", "apaitisi": 0.0}]
    if "bullets" not in st.session_state:
        st.session_state.bullets = [""]
    if "ek_override" not in st.session_state:
        st.session_state.ek_override = {}
    if "diap_captions" not in st.session_state:
        st.session_state.diap_captions = {}
    if "synim_captions" not in st.session_state:
        st.session_state.synim_captions = {}
    
    def add_row():
        st.session_state.damage_rows.append({"desc": "", "apaitisi": 0.0})
    
    def del_row(i):
        if len(st.session_state.damage_rows) > 1:
            st.session_state.damage_rows.pop(i)
            st.session_state.ek_override.pop(i, None)
    
    # =============================================================
    # UI – HEADER
    # =============================================================
    
    st.markdown("## 📋 Experts360 – Παραγωγή Έκθεσης Πραγματογνωμοσύνης")
    st.divider()
    
    template_bytes = None
    tp = Path("template_ekthesis.docx")
    if tp.exists():
        template_bytes = tp.read_bytes()
        st.success("✅ Template: `template_ekthesis.docx`")
    else:
        up = st.file_uploader("📂 Ανέβασε το `template_ekthesis.docx`", type=["docx"])
        if up:
            template_bytes = up.read()
    if not template_bytes:
        st.warning("⚠️ Βάλε το `template_ekthesis.docx` στον ίδιο φάκελο με το app.")
        st.stop()
    
    # =============================================================
    # ΣΤΟΙΧΕΙΑ ΖΗΜΙΑΣ & ΠΑΘΟΝΤΟΣ
    # =============================================================
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("## 🔢 Στοιχεία Ζημίας")
        arithmos   = st.text_input("Αριθμός Ζημίας *", placeholder="π.χ. 39150000")
        hm_epith   = st.date_input("Ημ. Επιθεώρησης *", value=date.today())
        hm_synt    = st.date_input("Ημ. Σύνταξης Έκθεσης", value=date.today())
        hm_anath   = st.date_input("Ημ. Ανάθεσης Εντολής", value=date.today())
        hm_zim     = st.date_input("Χρόνος Ζημίας", value=date.today())
    
    with col2:
        st.markdown("## 👤 Στοιχεία Παθόντος")
        pathon     = st.text_input("Παθών (ονοματεπώνυμο) *", placeholder="π.χ. ΠΑΠΑΔΟΠΟΥΛΟΥ ΜΑΡΙΑ")
        up_opsin   = st.text_input("Υπ' όψιν", placeholder="π.χ. κο. Παπαδόπουλο")
        asfalismen = st.text_input("Ασφαλισμένος / Αρ. Πινακίδας", placeholder="π.χ. ΧΧΧ1234")
        asfalister = st.text_input("Ασφαλιστήριο Συμβόλαιο", placeholder="π.χ. 70000000")
        topothesia = st.text_input("Τοποθεσία Ζημίας *", placeholder="π.χ. Λ. ΒΟΥΛΙΑΓΜΕΝΗΣ 150 ΓΛΥΦΑΔΑ")
    
    st.markdown("")
    
    # =============================================================
    # ΠΕΡΙΓΡΑΦΕΣ & ΦΩΤΟ
    # =============================================================
    
    st.markdown("## ✍️ Περιγραφές & Φωτογραφίες")
    thema = st.text_input("Θέμα Ζημίας", placeholder="π.χ. γυάλινη πόρτα φαρμακείου")
    
    col3, col4 = st.columns(2)
    
    with col3:
        eisagogi = st.text_area(
            "Εισαγωγική Παράγραφος",
            value=(
                "επικοινωνήσαμε με την παθούσα, ώστε να ενημερωθούμε για το ακριβές "
                "σημείο ζημιάς όπου και την ίδια ημέρα μετέβη ώστε να διαπιστώσουμε "
                "τις ζημιές που είχαν προκληθεί από το ασφαλισμένο όχημα υπ' αριθμόν "
                "κυκλοφορίας ΧΧΧ-1234"
            ),
            height=110,
        )
    
        st.markdown("**Γενικές Πληροφορίες – Στοιχεία Παθόντος**")
        genikes = st.text_area(
            "genikes", label_visibility="collapsed",
            value="",
            placeholder="π.χ. Επί της Λεωφόρου Βουλιαγμένης 150 στη Γλυφάδα η παθούσα διατηρεί φαρμακείο όπου η είσοδος αποτελείται από γυάλινη πόρτα με πλαίσιο αλουμινίου.",
            height=90,
        )
    
        st.markdown("📍 **Google Maps φωτό** *(προαιρετικό — εμφανίζεται κάτω από Γενικές Πληροφορίες)*")
        maps_file = st.file_uploader("maps", type=["jpg","jpeg","png"],
            label_visibility="collapsed", key="maps_up")
        if maps_file:
            st.image(maps_file, use_container_width=True)
    
    with col4:
        eisagogi_diap = st.text_input(
            "Εισαγωγή Διαπιστώσεων",
            value="Κατά την επιθεώρησή μας στο σημείο ζημίας, διαπιστώσαμε τα κάτωθι:",
        )
    
        st.markdown("**Διαπιστώσεις (bullets)**")
        for bi in range(len(st.session_state.bullets)):
            b1, b2 = st.columns([10, 1])
            with b1:
                st.session_state.bullets[bi] = st.text_input(
                    f"b{bi}", st.session_state.bullets[bi],
                    label_visibility="collapsed", key=f"bul_{bi}",
                    placeholder=f"π.χ. Σπάσιμο κρυστάλλου στη γυάλινη πόρτα εισόδου",
                )
            with b2:
                if st.button("✕", key=f"xb_{bi}") and len(st.session_state.bullets) > 1:
                    st.session_state.bullets.pop(bi)
                    st.rerun()
        if st.button("➕ Bullet"):
            st.session_state.bullets.append("")
            st.rerun()
    
        protasi = st.text_area(
            "Πρόταση Αποκατάστασης",
            value="Για την αποκατάσταση των προκληθέντων ζημιών θα πρέπει να αντικατασταθεί ",
            height=85,
        )
        prosopiki = st.text_area(
            "Προσωπική Άποψη *(προαιρετικό)*",
            placeholder="Συμπληρώνεται μόνο αν χρειάζεται...",
            height=60,
        )
    
    # Φωτό Διαπιστώσεων
    st.markdown("### 📸 Φωτό Διαπιστώσεων *(1–3, μπαίνουν κάτω από Πρόταση Αποκατάστασης)*")
    diap_files = st.file_uploader("diap", type=["jpg","jpeg","png"],
        accept_multiple_files=True, key="diap_up", label_visibility="collapsed")
    if diap_files:
        if len(diap_files) > 3:
            st.warning("⚠️ Μέγιστο 3 φωτό. Θα χρησιμοποιηθούν οι πρώτες 3.")
            diap_files = diap_files[:3]
        for i, f in enumerate(diap_files):
            c1, c2 = st.columns([2, 3])
            with c1:
                st.image(f, caption=f"Διαπίστωση {i+1}", use_container_width=True)
            with c2:
                st.session_state.diap_captions[i] = st.text_input(
                    f"Λεζάντα φωτό {i+1}",
                    value=st.session_state.diap_captions.get(i, ""),
                    placeholder="π.χ. Σπάσιμο κρυστάλλου εισόδου",
                    key=f"dcap_{i}",
                )
    
    apaitisi_par = st.text_area(
        "Παράγραφος Απαίτησης",
        value=(
            "Για την εξεταζόμενη ζημία, ο παθών μας απέστειλε την οικονομική προσφορά, "
            "αντί απαιτήσεως, για την αποκατάσταση των ζημιών, έναντι του χρηματικού ποσού "
            "των        €  (συμπεριλαμβανομένου του Φ.Π.Α.), το οποίο αναλύεται λεπτομερώς "
            "στο κεφάλαιο της εκτίμησης που ακολουθεί."
        ),
        height=100,
    )
    
    st.markdown("")
    
    # =============================================================
    # ΠΙΝΑΚΑΣ ΚΟΣΤΟΥΣ
    # =============================================================
    
    st.markdown("## 💶 Πίνακας Κόστους")
    st.caption("Η εκτίμηση ισούται με την απαίτηση by default — άλλαξέ την αν χρειάζεται. ΦΠΑ & Σύνολα αυτόματα.")
    
    hc = st.columns([5, 2, 2, 1])
    hc[0].markdown("**Περιγραφή**")
    hc[1].markdown("**Απαίτηση (€)**")
    hc[2].markdown("**Εκτίμηση (€)**")
    
    total_ap = 0.0
    total_ek = 0.0
    rows_data = []
    
    for i, row in enumerate(st.session_state.damage_rows):
        rc = st.columns([5, 2, 2, 1])
        with rc[0]:
            desc = st.text_input(f"d{i}", value=row["desc"],
                label_visibility="collapsed", key=f"desc_{i}",
                placeholder=f"π.χ. Αντικατάσταση γυάλινης πόρτας εισόδου")
        with rc[1]:
            ap = st.number_input("a", value=float(row["apaitisi"]),
                min_value=0.0, step=10.0, format="%.2f",
                label_visibility="collapsed", key=f"ap_{i}")
        ek_def = st.session_state.ek_override.get(i, ap)
        with rc[2]:
            ek = st.number_input("e", value=float(ek_def),
                min_value=0.0, step=10.0, format="%.2f",
                label_visibility="collapsed", key=f"ek_{i}")
            st.session_state.ek_override[i] = ek
        with rc[3]:
            if st.button("🗑️", key=f"del_{i}"):
                del_row(i)
                st.rerun()
        st.session_state.damage_rows[i] = {"desc": desc, "apaitisi": ap}
        total_ap += ap
        total_ek += ek
        rows_data.append({"desc": desc, "apaitisi": ap, "ektimisi": ek})
    
    st.button("➕ Προσθήκη γραμμής", on_click=add_row)
    
    fpa_ap = total_ap * 0.24
    fpa_ek = total_ek * 0.24
    tel_ap = total_ap + fpa_ap
    tel_ek = total_ek + fpa_ek
    
    st.markdown(f"""
    <div class="total-box">
    <table style="width:100%;border-collapse:collapse;">
    <tr><td style="width:55%"></td>
        <th style="text-align:right;padding:4px 16px;">Απαίτηση</th>
        <th style="text-align:right;padding:4px 16px;">Εκτίμηση</th></tr>
    <tr><td>Μερικό σύνολο</td>
        <td style="text-align:right;padding:4px 16px;">{fmt_euro(total_ap)}</td>
        <td style="text-align:right;padding:4px 16px;">{fmt_euro(total_ek)}</td></tr>
    <tr><td>ΦΠΑ 24%</td>
        <td style="text-align:right;padding:4px 16px;">{fmt_euro(fpa_ap)}</td>
        <td style="text-align:right;padding:4px 16px;">{fmt_euro(fpa_ek)}</td></tr>
    <tr style="font-weight:bold;border-top:2px solid #388e3c;">
        <td>ΣΥΝΟΛΟ</td>
        <td style="text-align:right;padding:4px 16px;">{fmt_euro(tel_ap)}</td>
        <td style="text-align:right;padding:4px 16px;">{fmt_euro(tel_ek)}</td></tr>
    </table>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("")
    
    # =============================================================
    # ΦΩΤΟ ΣΥΝΝΗΜΕΝΩΝ
    # =============================================================
    
    st.markdown("## 🖼️ Φωτό Συννημένων *(μπαίνουν κάτω από ΦΩΤΟΓΡΑΦΙΚΟ ΥΛΙΚΟ)*")
    synim_files = st.file_uploader("synim", type=["jpg","jpeg","png"],
        accept_multiple_files=True, key="synim_up", label_visibility="collapsed")
    if synim_files:
        for i, f in enumerate(synim_files):
            c1, c2 = st.columns([2, 3])
            with c1:
                st.image(f, caption=f"Φωτό {i+1}", use_container_width=True)
            with c2:
                st.session_state.synim_captions[i] = st.text_input(
                    f"Λεζάντα φωτό {i+1}",
                    value=st.session_state.synim_captions.get(i, ""),
                    placeholder="π.χ. Άποψη ζημίας από αριστερά",
                    key=f"scap_{i}",
                )
    
    st.markdown("")
    
    # =============================================================
    # ΠΑΡΑΓΩΓΗ
    # =============================================================
    
    st.markdown("## 📄 Παραγωγή Εκθέσεως")
    
    fn_default = (
        f"ekthesi_{arithmos}_{fmt_short(hm_synt).replace('.','')}.docx"
        if arithmos else "nea_ekthesi.docx"
    )
    out_name = st.text_input("Όνομα αρχείου εξόδου", value=fn_default)
    
    # Αποθήκευση στη βάση
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
        eid_k = st.session_state.get('current_ktirion_id')
        btn_lbl = f"💾 Ενημέρωση #{eid_k}" if eid_k else "💾 Αποθήκευση νέας"
        if st.button(btn_lbl, type="primary", use_container_width=True, key="ktir_save_btn"):
            save_data_k = {
                "arithmos_zimias": arithmos,
                "hm_epitheorisis": str(hm_epith),
                "hm_syntaxis":     str(hm_synt),
                "hm_anathesis":    str(hm_anath),
                "hm_zimias":       str(hm_zim),
                "pathon":          pathon,
                "up_opsin":        up_opsin,
                "asfalismenos":    asfalismen,
                "asfalisterio":    asfalister,
                "topothesia":      topothesia,
                "thema":           thema,
                "eisagogi":        eisagogi,
                "genikes":         genikes,
                "eisagogi_diap":   eisagogi_diap,
                "bullets":         "\n".join(st.session_state.bullets),
                "protasi":         protasi,
                "prosopiki":       prosopiki,
                "apaitisi_par":    apaitisi_par,
                "total_apaitisi":  total_ap,
                "total_ektimisi":  total_ek,
                "asfalistiki":     st.session_state.get('asfalistiki',''),
                "onomateponymo":   save_user_k,
                "status":          save_status_k,
            }
            new_id_k, err_k = save_ekthesi_ktiriou(
                data=save_data_k, grammes=rows_data,
                user_name=save_user_k, ekthesi_id=eid_k
            )
            if err_k:
                st.error(f"❌ {err_k}")
            else:
                st.session_state['current_ktirion_id'] = new_id_k
                st.success(f"✅ Αποθηκεύτηκε #{new_id_k}")

    st.markdown("---")
    st.subheader("📄 Παραγωγή Εκθέσεως")
    if st.button("🖨️ Δημιουργία Εκθέσεως", type="primary", use_container_width=True):
    
        errs = []
        if not arithmos:   errs.append("Αριθμός Ζημίας")
        if not pathon:     errs.append("Παθών")
        if not topothesia: errs.append("Τοποθεσία Ζημίας")
        if all(r["desc"] == "" for r in rows_data): errs.append("Τουλάχιστον μία γραμμή ζημίας")
    
        if errs:
            st.error(f"⚠️ Συμπλήρωσε: {', '.join(errs)}")
        else:
            MAX = 10
            tbl = {}
            for idx in range(MAX):
                n = idx + 1
                if idx < len(rows_data):
                    r = rows_data[idx]
                    tbl[f"{{{{ΕΙΔΟΣ_{n}}}}}"]    = r["desc"]
                    tbl[f"{{{{ΑΠΑΙΤΗΣΗ_{n}}}}}"] = fmt_euro(r["apaitisi"])
                    tbl[f"{{{{ΕΚΤΙΜΗΣΗ_{n}}}}}"] = fmt_euro(r["ektimisi"])
                else:
                    tbl[f"{{{{ΕΙΔΟΣ_{n}}}}}"]    = ""
                    tbl[f"{{{{ΑΠΑΙΤΗΣΗ_{n}}}}}"] = ""
                    tbl[f"{{{{ΕΚΤΙΜΗΣΗ_{n}}}}}"] = ""
    
            bullets_text = "\n".join(b for b in st.session_state.bullets if b.strip())
    
            data = {
                "{{ΑΡΙΘΜΟΣ_ΖΗΜΙΑΣ}}":        arithmos,
                "{{ΗΜ_ΕΠΙΘΕΩΡΗΣΗΣ}}":        fmt_short(hm_epith),
                "{{ΗΜΕΡΟΜΗΝΙΑ_ΣΥΝΤΑΞΗΣ}}":   fmt_long(hm_synt),
                "{{ΘΕΜΑ_ΖΗΜΙΑΣ}}":            thema,
                "{{ΥΠ_ΟΨΙΝ}}":                up_opsin,
                "{{ΑΣΦΑΛΙΣΜΕΝΟΣ}}":           asfalismen,
                "{{ΑΣΦΑΛΙΣΤΗΡΙΟ}}":           asfalister,
                "{{ΠΑΘΩΝ}}":                  pathon,
                "{{ΤΟΠΟΘΕΣΙΑ_ΖΗΜΙΑΣ}}":      topothesia,
                "{{ΧΡΟΝΟΣ_ΖΗΜΙΑΣ}}":         fmt_short(hm_zim),
                "{{ΗΜ_ΑΝΑΘΕΣΗΣ}}":           fmt_short(hm_anath),
                "{{ΑΠΑΙΤΗΣΗ_ΣΥΝΟΛΟ}}":       fmt_euro(total_ap),
                "{{ΕΚΤΙΜΗΣΗ_ΣΥΝΟΛΟ}}":       fmt_euro(total_ek),
                "{{ΕΙΣΑΓΩΓΗ_ΠΕΡΙΓΡΑΦΗ}}":    eisagogi,
                "{{ΓΕΝΙΚΕΣ_ΠΛΗΡΟΦΟΡΙΕΣ}}":   genikes,
                "{{ΕΙΣΑΓΩΓΗ_ΔΙΑΠΙΣΤΩΣΕΩΝ}}": eisagogi_diap,
                "{{ΔΙΑΠΙΣΤΩΣΕΙΣ_BULLETS}}":  bullets_text,
                "{{ΠΡΟΤΑΣΗ_ΑΠΟΚΑΤΑΣΤΑΣΗΣ}}": protasi,
                "{{ΠΡΟΣΩΠΙΚΗ_ΑΠΟΨΗ}}":       prosopiki,
                "{{ΑΠΑΙΤΗΣΗ_ΠΕΡΙΓΡΑΦΗ}}":    apaitisi_par,
                **tbl,
                "{{ΜΕΡΙΚΟ_ΣΥΝΟΛΟ_ΑΠΑΙΤΗΣΗ}}": fmt_euro(total_ap),
                "{{ΜΕΡΙΚΟ_ΣΥΝΟΛΟ_ΕΚΤΙΜΗΣΗ}}": fmt_euro(total_ek),
                "{{ΦΠΑ_ΑΠΑΙΤΗΣΗ}}":           fmt_euro(fpa_ap),
                "{{ΦΠΑ_ΕΚΤΙΜΗΣΗ}}":           fmt_euro(fpa_ek),
                "{{ΣΥΝΟΛΟ_ΑΠΑΙΤΗΣΗ}}":        fmt_euro(tel_ap),
                "{{ΣΥΝΟΛΟ_ΕΚΤΙΜΗΣΗ}}":        fmt_euro(tel_ek),
            }
    
            maps_b  = to_jpeg(maps_file) if maps_file else None
            diap_b  = [(to_jpeg(f), st.session_state.diap_captions.get(i, ""))
                       for i, f in enumerate(diap_files or [])]
            synim_b = [(to_jpeg(f), st.session_state.synim_captions.get(i, ""))
                       for i, f in enumerate(synim_files or [])]
    
            with st.spinner("Παράγεται η έκθεση..."):
                try:
                    result = fill_template(template_bytes, data, maps_b, diap_b, synim_b)
                    st.success("✅ Η έκθεση δημιουργήθηκε!")
                    # Debug: έλεγξε αν τα placeholders αντικαταστάθηκαν
                    import zipfile as zf2
                    with zf2.ZipFile(io.BytesIO(result)) as zcheck:
                        hcheck = zcheck.read("word/header1.xml").decode("utf-8")
                    if "{{ΑΡΙΘΜΟΣ_ΖΗΜΙΑΣ}}" in hcheck:
                        st.warning("⚠️ DEBUG: Το placeholder header ΔΕΝ αντικαταστάθηκε!")
                    else:
                        st.info(f"✅ DEBUG: Header ΟΚ")
                    st.download_button(
                        label=f"⬇️ Κατέβασε: {out_name}",
                        data=result,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"❌ Σφάλμα: {e}")
                    import traceback
                    st.code(traceback.format_exc())
    
