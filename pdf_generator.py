"""
pdf_generator.py - Δημιουργία PDF με ελληνικά (DejaVu font)
"""
import io, os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Table, TableStyle, Image, PageBreak)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Εγγραφή ελληνικών γραμματοσειρών
FONT_PATHS = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/System/Library/Fonts/Helvetica.ttc',
    '/Library/Fonts/Arial.ttf',
]
FONTB_PATHS = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    '/System/Library/Fonts/Helvetica.ttc',
    '/Library/Fonts/Arial Bold.ttf',
]

_font_reg = False

def register_fonts():
    global _font_reg
    if _font_reg:
        return

    # Ψάχνει στον ίδιο φάκελο με το pdf_generator.py πρώτα
    _dir = os.path.dirname(os.path.abspath(__file__))
    search_paths_reg = [
        os.path.join(_dir, "DejaVuSans.ttf"),
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
    ]
    search_paths_bold = [
        os.path.join(_dir, "DejaVuSans-Bold.ttf"),
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
        '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
    ]

    found_reg  = next((p for p in search_paths_reg  if os.path.exists(p)), None)
    found_bold = next((p for p in search_paths_bold if os.path.exists(p)), None)

    if found_reg:
        pdfmetrics.registerFont(TTFont('GR', found_reg))
        pdfmetrics.registerFont(TTFont('GR-Bold', found_bold or found_reg))
    else:
        raise Exception("Δεν βρέθηκε γραμματοσειρά. Βάλε DejaVuSans.ttf στον φάκελο της εφαρμογής.")

    _font_reg = True

# Χρώματα
C_BLUE  = colors.HexColor('#1F3864')
C_LBLUE = colors.HexColor('#D6E4F0')
C_DGRAY = colors.HexColor('#BFBFBF')
C_LGRAY = colors.HexColor('#F2F2F2')
C_WHITE = colors.white
C_BLACK = colors.black

def S(name, **kw):
    register_fonts()
    d = dict(fontName='GR', fontSize=8, textColor=C_BLACK, leading=10)
    d.update(kw)
    return ParagraphStyle(name, **d)

def P(txt, style):
    return Paragraph(str(txt or '').replace('\n','<br/>'), style)

def fmt(v):
    try:
        f = float(v or 0)
        if f == 0: return ''
        return f"{f:,.2f}".replace(',','X').replace('.',',').replace('X','.')
    except: return ''

def fmt0(v):
    try:
        f = float(v or 0)
        return f"{f:,.2f}".replace(',','X').replace('.',',').replace('X','.')
    except: return '0,00'


def generate_pdf(data, parts, works, visits, photo_files,
                 photo_captions, paratiriseis, onomateponymo,
                 asfalistiki='INTERLIFE'):
    register_fonts()
    buf = io.BytesIO()
    W, H = A4
    ML = MR = 1.8*cm
    MT = MB = 1.5*cm
    CW = W - ML - MR

    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=ML, rightMargin=MR,
                            topMargin=MT, bottomMargin=MB)
    story = []

    # Styles
    s_title = S('t',  fontName='GR-Bold', fontSize=13, alignment=TA_CENTER)
    s_lbl   = S('l',  fontName='GR-Bold', fontSize=8)
    s_val   = S('v',  fontName='GR',      fontSize=8)
    s_sm    = S('s',  fontName='GR',      fontSize=8)
    s_bsm   = S('bs', fontName='GR-Bold', fontSize=8)
    s_hdrl  = S('hl', fontName='GR-Bold', fontSize=8, textColor=C_WHITE)
    s_hdr   = S('h',  fontName='GR-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)
    s_r     = S('r',  fontName='GR',      fontSize=8, alignment=TA_RIGHT)
    s_br    = S('br', fontName='GR-Bold', fontSize=8, alignment=TA_RIGHT)
    s_bc    = S('bc', fontName='GR-Bold', fontSize=8, alignment=TA_CENTER)
    s_white = S('w',  fontName='GR-Bold', fontSize=10, textColor=C_WHITE, alignment=TA_CENTER)
    s_sec   = S('sec',fontName='GR-Bold', fontSize=8)
    s_notes = S('n',  fontName='GR',      fontSize=8, leading=12)
    s_cap   = S('cp', fontName='GR',      fontSize=8, alignment=TA_CENTER,
                textColor=colors.HexColor('#555555'))

    # ── HEADER ──────────────────────────────────────────────
    # Λογότυπα ανάλογα με ασφαλιστική - από τοπικά αρχεία
    _dir = os.path.dirname(os.path.abspath(__file__))

    if asfalistiki == 'INTERLIFE':
        gnomon_path = os.path.join(_dir, "logo_gnomon.png")
        tuv_path    = os.path.join(_dir, "logo_tuv.jpg")
        if os.path.exists(gnomon_path) and os.path.exists(tuv_path):
            g_img = Image(gnomon_path, width=4*cm, height=1.5*cm)
            t_img = Image(tuv_path,    width=1.5*cm, height=1.5*cm)
            hdr_data = [[g_img, t_img]]
            hdr_cols = [CW-2*cm, 2*cm]
        elif os.path.exists(gnomon_path):
            hdr_data = [[Image(gnomon_path, width=4*cm, height=1.5*cm), '']]
            hdr_cols = [CW-2*cm, 2*cm]
        else:
            hdr_data = [[P('GNOMON EXPERTS', s_title), P('TUV', s_bc)]]
            hdr_cols = [CW-2*cm, 2*cm]
    else:
        logo_path = os.path.join(_dir, "logo_experts360.png")
        if os.path.exists(logo_path):
            hdr_data = [[Image(logo_path, width=4*cm, height=1.4*cm), '']]
        else:
            hdr_data = [[P('Experts360', s_title), '']]
        hdr_cols = [CW-2*cm, 2*cm]

    ht = Table(hdr_data, colWidths=hdr_cols)
    ht.setStyle(TableStyle([
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('ALIGN',(1,0),(1,0),'RIGHT'),
        ('BOTTOMPADDING',(0,0),(-1,-1),2),
    ]))
    story.append(ht)
    story.append(Spacer(1,2*mm))
    story.append(P('ΕΚΘΕΣΗ ΠΡΑΓΜΑΤΟΓΝΩΜΟΣΥΝΗΣ', s_title))
    story.append(Spacer(1,3*mm))

    # Section title
    def section(txt):
        t = Table([[P(txt, s_sec)]], colWidths=[CW])
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1), C_LBLUE),
            ('TOPPADDING',(0,0),(-1,-1),3),
            ('BOTTOMPADDING',(0,0),(-1,-1),3),
            ('LEFTPADDING',(0,0),(-1,-1),5),
            ('BOX',(0,0),(-1,-1),0.5,C_DGRAY),
        ]))
        return t

    # Info row styles
    TS = TableStyle([
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),2),
        ('BOTTOMPADDING',(0,0),(-1,-1),2),
        ('LEFTPADDING',(0,0),(-1,-1),4),
        ('RIGHTPADDING',(0,0),(-1,-1),4),
        ('LINEBELOW',(0,0),(-1,-1),0.3,C_DGRAY),
    ])

    def kv2(l1,v1,l2='',v2=''):
        if l2:
            t=Table([[P(l1,s_lbl),P(v1,s_val),P(l2,s_lbl),P(v2,s_val)]],
                    colWidths=[4*cm,5*cm,3.5*cm,CW-12.5*cm])
        else:
            t=Table([[P(l1,s_lbl),P(v1,s_val)]],
                    colWidths=[4*cm,CW-4*cm])
        t.setStyle(TS)
        return t

    def kv3(l1,v1,l2,v2,l3,v3):
        t=Table([[P(l1,s_lbl),P(v1,s_val),P(l2,s_lbl),P(v2,s_val),P(l3,s_lbl),P(v3,s_val)]],
                colWidths=[2.5*cm,3.5*cm,2*cm,3.5*cm,2.5*cm,CW-14*cm])
        t.setStyle(TS)
        return t

    # ΣΤΟΙΧΕΙΑ ΑΤΥΧΗΜΑΤΟΣ
    story.append(section('ΣΤΟΙΧΕΙΑ ΑΤΥΧΗΜΑΤΟΣ'))
    story.append(kv2('Αρ. Ζημίας:',       data.get('ar_zimias',''),
                      'Ημ/να Εντολής:',   data.get('hm_entolhs','')))
    story.append(kv2('Ημ/νία Ατυχήματος:', data.get('hm_atyxhmatos','')))
    story.append(kv2('Τόπος Ατυχήματος:',  data.get('topos_atyxhmatos','')))
    story.append(Spacer(1,2*mm))

    # ΣΤΟΙΧΕΙΑ ΟΧΗΜΑΤΟΣ
    story.append(section('ΣΤΟΙΧΕΙΑ ΕΠΙΘΕΩΡΟΥΜΕΝΟΥ ΟΧΗΜΑΤΟΣ'))
    story.append(kv2('Ιδιοκτήτης:',           data.get('idioktitis','')))
    story.append(kv2('Αρ. Κυκλοφορίας:',      data.get('ar_kykloforias',''),
                      'Χρήση:',               data.get('xrisi','')))
    story.append(kv3('Μάρκα:',               data.get('marka',''),
                      'Κυβικά:',             data.get('kyvika',''),
                      'Μοντέλο:',            data.get('montelo','')))
    story.append(kv2('1η Άδεια Κυκλοφορίας:',data.get('proti_adeia',''),
                      'Ημ/νία ΚΤΕΟ:',        data.get('hm_kteo','')))
    story.append(kv2('Ένδειξη Χιλιομετρητή:',data.get('xiliometrites',''),
                      'Αξία:',
                      f"{data.get('axia','')} €" if data.get('axia') else ''))
    if data.get('ar_plaisiou'):
        story.append(kv2('Αρ. Πλαισίου (VIN):', data.get('ar_plaisiou','')))
    story.append(Spacer(1,2*mm))

    # ΕΠΙΘΕΩΡΗΣΗ
    story.append(section('ΣΤΟΙΧΕΙΑ ΕΠΙΘΕΩΡΗΣΗΣ'))
    ords = ['1ης','2ης','3ης','4ης']
    for i,v in enumerate(visits):
        o = ords[i] if i<4 else f'{i+1}ης'
        story.append(kv2(f'Ημ/νία {o} Επίσκεψης:',
                         v.get('date_text','') or '',
                         'Τόπος:', v.get('place','') or ''))
    story.append(Spacer(1,2*mm))

    # ΠΙΝΑΚΕΣ
    cw = [CW-12*cm,1.8*cm,1.8*cm,1.8*cm,1.8*cm,1.8*cm,2.2*cm]

    def num(v): return P(fmt(v), s_r)
    def rs(row):
        return sum(float(row.get(k,0) or 0)
                   for k in ['price','fanop','vafeas','mixanikos','ilgos'])

    def make_table(rows_data, header_label):
        tbl = [[P(header_label,s_hdrl),
                P('Τιμή',s_hdr), P('Φανοπ.',s_hdr),
                P('Βαφέας',s_hdr), P('Μηχ/κός',s_hdr),
                P('Ηλ/γος',s_hdr), P('Σύνολα',s_hdr)]]
        for row in rows_data:
            t = rs(row)
            tbl.append([P(row['desc'],s_sm),
                        num(row.get('price')), num(row.get('fanop')),
                        num(row.get('vafeas')), num(row.get('mixanikos')),
                        num(row.get('ilgos')),
                        P(fmt(t) if t else '', s_r)])
        while len(tbl) < 11:
            tbl.append(['','','','','','',''])
        t = Table(tbl, colWidths=cw, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0), C_BLUE),
            ('TEXTCOLOR',(0,0),(-1,0), C_WHITE),
            ('FONTNAME',(0,0),(-1,0),'GR-Bold'),
            ('GRID',(0,0),(-1,-1),0.3,C_DGRAY),
            ('BOX',(0,0),(-1,-1),0.5,C_DGRAY),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1),2),
            ('BOTTOMPADDING',(0,0),(-1,-1),2),
            ('LEFTPADDING',(0,0),(-1,-1),3),
            ('RIGHTPADDING',(0,0),(-1,-1),3),
        ]))
        return t

    # Ανταλλακτικά
    act_p = [{'desc': f"ΑΝΤΙΚΑΤΑΣΤΑΣΗ {x.get('name','')}" +
                      (f" ({x['type']})" if x.get('type') else ''),
              **x}
             for x in parts if x.get('name','').strip()]
    story.append(make_table(act_p, 'Ανταλλακτικά'))
    story.append(Spacer(1,2*mm))

    # Εργασίες
    act_w = [{'desc': (w.get('type','') + ' ' + w.get('desc','')).strip(), **w}
             for w in works if w.get('type','').strip() or w.get('desc','').strip()]
    story.append(make_table(act_w, 'Εργασίες'))
    story.append(Spacer(1,2*mm))

    # Σύνολα
    all_r = act_p + act_w
    def cs(k): return sum(float(r.get(k,0) or 0) for r in all_r)
    sp,sf,sv,sm,si = cs('price'),cs('fanop'),cs('vafeas'),cs('mixanikos'),cs('ilgos')
    st_ = sp+sf+sv+sm+si
    grand = st_ * 1.24

    def srow(lbl,*vals):
        return [P(lbl,s_bsm)]+[P(fmt0(v),s_br) for v in vals]

    sum_tbl = [
        srow('Σύνολα χωρίς Φ.Π.Α.', sp,sf,sv,sm,si,st_),
        srow('Φ.Π.Α. 24%', sp*.24,sf*.24,sv*.24,sm*.24,si*.24,st_*.24),
        srow('Σύνολα τιμών με Φ.Π.Α.', sp*1.24,sf*1.24,sv*1.24,sm*1.24,si*1.24,grand),
    ]
    st = Table(sum_tbl, colWidths=cw)
    st.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),'GR-Bold'),
        ('GRID',(0,0),(-1,-1),0.3,C_DGRAY),
        ('BOX',(0,0),(-1,-1),0.5,C_DGRAY),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),2),
        ('BOTTOMPADDING',(0,0),(-1,-1),2),
        ('LEFTPADDING',(0,0),(-1,-1),3),
        ('RIGHTPADDING',(0,0),(-1,-1),3),
    ]))
    story.append(st)
    story.append(Spacer(1,3*mm))

    # Γενικό Σύνολο
    gt = Table([[P('Γενικό Σύνολο Ζημίας:',
                   S('gl',fontName='GR-Bold',fontSize=10)),
                 P(f'{fmt0(grand)} €', s_white)]],
               colWidths=[CW-3.2*cm, 3.2*cm])
    gt.setStyle(TableStyle([
        ('BACKGROUND',(1,0),(1,0),C_BLUE),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),5),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('LEFTPADDING',(0,0),(-1,-1),5),
        ('BOX',(0,0),(-1,-1),0.5,C_DGRAY),
    ]))
    story.append(gt)
    story.append(Spacer(1,4*mm))

    # ΠΑΡΑΤΗΡΗΣΕΙΣ
    story.append(P('ΠΑΡΑΤΗΡΗΣΕΙΣ:', s_lbl))
    story.append(Spacer(1,1*mm))
    par_tbl = Table([[P(paratiriseis.strip() if paratiriseis else '', s_notes)]],
                    colWidths=[CW])
    par_tbl.setStyle(TableStyle([
        ('BOX',(0,0),(-1,-1),0.5,C_DGRAY),
        ('TOPPADDING',(0,0),(-1,-1),5),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('LEFTPADDING',(0,0),(-1,-1),5),
        ('RIGHTPADDING',(0,0),(-1,-1),5),
        ('MINROWHEIGHT',(0,0),(-1,-1),2.5*cm),
    ]))
    story.append(par_tbl)
    story.append(Spacer(1,5*mm))

    # ΥΠΟΓΡΑΦΗ
    sign_text = 'ΓΙΑ ΤΗΝ  GNOMON EXPERTS Α.Ε.' if asfalistiki == 'INTERLIFE' else 'ΓΙΑ ΤΗΝ  EXPERTS360'
    story.append(P(sign_text, s_lbl))
    story.append(Spacer(1,8*mm))
    ul = Table([[P('_'*45, s_val)]], colWidths=[CW/2])
    story.append(ul)
    if onomateponymo:
        story.append(P(onomateponymo, s_val))

    # ΦΩΤΟΓΡΑΦΙΕΣ
    if photo_files:
        story.append(PageBreak())
        story.append(P('ΦΩΤΟΓΡΑΦΙΚΟ ΥΛΙΚΟ', s_title))
        story.append(Spacer(1,5*mm))
        img_h = (A4[1]-4*cm)/2 - 1.5*cm
        for i,pf in enumerate(photo_files):
            try:
                pf.seek(0)
                # Συμπίεση εικόνας πριν το PDF
                from PIL import Image as PILImg
                raw = pf.read()
                pil = PILImg.open(io.BytesIO(raw))
                # Max 1200px πλάτος
                if pil.width > 1200:
                    ratio = 1200 / pil.width
                    pil = pil.resize((1200, int(pil.height * ratio)), PILImg.LANCZOS)
                if pil.mode in ('RGBA','P'):
                    pil = pil.convert('RGB')
                buf = io.BytesIO()
                pil.save(buf, format='JPEG', quality=60, optimize=True)
                buf.seek(0)
                img = Image(buf, width=CW, height=img_h)
                img.hAlign = 'CENTER'
                story.append(img)
                cap = photo_captions[i] if i<len(photo_captions) else ''
                if cap:
                    story.append(Spacer(1,2*mm))
                    story.append(P(cap, s_cap))
                story.append(Spacer(1,5*mm))
                if (i+1)%2==0 and i<len(photo_files)-1:
                    story.append(PageBreak())
            except Exception as e:
                story.append(P(f'[Σφάλμα φωτογραφίας]', s_sm))

    doc.build(story)
    buf.seek(0)
    return buf.read()
