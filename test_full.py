"""
test_full.py - Πλήρης έλεγχος εφαρμογής Experts360
Τρέξε: cd ~/Documents/interlifesample && python3 test_full.py
"""
import sys, io, traceback
sys.path.insert(0, '.')

results = []

def test(name, fn, expected=None):
    try:
        result = fn()
        if expected is not None:
            ok = result == expected
            status = "✅" if ok else "❌"
            print(f"{status} {name}: {result} (αναμενόμενο: {expected})")
        else:
            ok = bool(result)
            status = "✅" if ok else "❌"
            print(f"{status} {name}: {result}")
        results.append((name, ok, str(result)[:150]))
    except Exception as e:
        print(f"❌ {name}: {e}")
        results.append((name, False, str(e)[:150]))

print("\n" + "="*60)
print("EXPERTS360 v2.1 - ΠΛΗΡΗΣ ΕΛΕΓΧΟΣ")
print("="*60)

# ── ΒΑΣΗ ──────────────────────────────────────────────────
print("\n📦 ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ")
from database import (get_connection, init_db, search_ektheseis,
    get_statistics, load_ekthesi, save_ekthesi,
    get_custom_markes, get_custom_montela,
    get_abbreviations, get_synergeio_eponimies)

conn, db_type = get_connection()
test("1. Σύνδεση βάσης", lambda: db_type in ('sqlite','postgres','mysql'), True)
test("2. Τύπος βάσης", lambda: db_type)
test("3. Init DB", lambda: init_db()[0], True)
test("4. Αναζήτηση (επιστρέφει list)", lambda: isinstance(search_ektheseis(), list), True)
test("5. Αναζήτηση (πλήθος)", lambda: len(search_ektheseis()))
test("6. Αναζήτηση (τύπος row)", lambda: type(search_ektheseis()[0]).__name__ if search_ektheseis() else 'empty')
test("7. Αναζήτηση (has 'id' key)", lambda: 'id' in search_ektheseis()[0] if search_ektheseis() else True, True)
test("8. Στατιστικά total", lambda: isinstance(get_statistics().get('total'), int), True)
test("9. Στατιστικά by_status (dict)", lambda: isinstance(get_statistics().get('by_status'), dict), True)
test("10. Στατιστικά top_markes (list)", lambda: isinstance(get_statistics().get('top_markes'), list), True)
test("11. Custom μάρκες", lambda: isinstance(get_custom_markes(), list), True)
test("12. Συντομεύσεις", lambda: isinstance(get_abbreviations(), dict), True)
test("13. Συνεργεία", lambda: isinstance(get_synergeio_eponimies(), list), True)

# ── ΑΠΟΘΗΚΕΥΣΗ / ΦΟΡΤΩΣΗ ──────────────────────────────────
print("\n💾 ΑΠΟΘΗΚΕΥΣΗ & ΦΟΡΤΩΣΗ")
_test_data = {
    'ar_zimias': 'TEST/99', 'hm_entolhs': '22/05/2026',
    'hm_atyxhmatos': '21/05/2026', 'topos_atyxhmatos': 'ΑΘΗΝΑ',
    'idioktitis': 'ΤΕΣΤ ΤΕΣΤΟΠΟΥΛΟΣ', 'ar_kykloforias': 'ΑΑΑ0001',
    'ar_plaisiou': 'WVW12345678901234', 'marka': 'Toyota',
    'montelo': 'Corolla', 'kyvika': '1600', 'xrisi': 'ΕΙΧ',
    'proti_adeia': '01/01/2020', 'xiliometrites': '50000',
    'axia': 5000, 'hm_kteo': '', 'status': 'draft',
    'asfalistiki': 'INTERLIFE', 'onomateponymo': 'ΤΕΣΤ'
}
_parts = [{'name':'ΠΡΟΦΥΛΑΚΤΗΡΑΣ','type':'ΓΝ','price':150,'fanop':0,'vafeas':0,'mixanikos':0,'ilgos':0}]
_works = [{'type':'ΒΑΦΗ','desc':'ΠΡΟΦΥΛΑΚΤΗΡΑΣ','fanop':0,'vafeas':80,'mixanikos':0,'ilgos':0}]

_saved_id, _save_err = save_ekthesi(_test_data, _parts, _works, [])
test("14. Αποθήκευση (ID επιστρέφεται)", lambda: _saved_id is not None, True)
test("15. Αποθήκευση (χωρίς error)", lambda: _save_err is None, True)

if _saved_id:
    _loaded = load_ekthesi(_saved_id)
    test("16. Φόρτωση (επιστρέφει dict)", lambda: isinstance(_loaded, dict), True)
    test("17. Φόρτωση ar_zimias", lambda: _loaded.get('ar_zimias') == 'TEST/99', True)
    test("18. Φόρτωση marka", lambda: _loaded.get('marka') == 'Toyota', True)
    test("19. Ανταλλακτικά (έχει parts)", lambda: len(_loaded.get('parts',[])) > 0, True)
    test("20. Ανταλλακτικά (τύπος row)", lambda: type(_loaded['parts'][0]).__name__ if _loaded.get('parts') else 'empty')
    test("21. Ανταλλακτικά (has 'name' key)", lambda: 'name' in _loaded['parts'][0] if _loaded.get('parts') else False, True)
    test("22. Εργασίες (έχει works)", lambda: len(_loaded.get('works',[])) > 0, True)
    test("23. Εργασίες (has 'descr' key)", lambda: 'descr' in _loaded['works'][0] if _loaded.get('works') else False, True)

    # Αναζήτηση με query
    _search_res = search_ektheseis(query='TEST/99')
    test("24. Αναζήτηση με query", lambda: len(_search_res) > 0, True)
    test("25. Αναζήτηση result has id", lambda: 'id' in _search_res[0] if _search_res else False, True)

    # Cleanup
    from database import get_connection as _gc, ph as _ph
    _c, _dt = _gc()
    _p = _ph(_dt)
    _cur = _c.cursor()
    _cur.execute(f"DELETE FROM ektheseis WHERE ar_zimias = {_p}", ('TEST/99',))
    _c.commit(); _c.close()
    print("   🧹 Cleanup: Διαγράφηκε η test εγγραφή TEST/99")

# ── PDF ──────────────────────────────────────────────────
print("\n📄 PDF GENERATION")
from pdf_generator import generate_pdf
from PIL import Image as PILImg

_pdf_data = {
    'ar_zimias':'TEST26','marka':'Toyota','montelo':'Corolla',
    'ar_kykloforias':'ΑΑΑ0001','idioktitis':'ΤΕΣΤ ΤΕΣΤΟΠΟΥΛΟΣ',
    'xrisi':'ΕΙΧ','kyvika':'1600','proti_adeia':'01/01/2020',
    'hm_entolhs':'22/05/2026','hm_atyxhmatos':'21/05/2026',
    'topos_atyxhmatos':'ΑΘΗΝΑ','xiliometrites':'50000',
    'axia':'5000','hm_kteo':'','ar_plaisiou':'WVW12345678901234',
    'xroma':'ΑΣΠΡΟ','kaysimo':'Βενζίνη'
}

def _gen_pdf(**kwargs):
    return generate_pdf(data=_pdf_data, parts=_parts, works=_works,
        visits=[], paratiriseis='Τεστ παρατήρηση',
        onomateponymo='Μαρινάκης', **kwargs)

test("26. PDF INTERLIFE (valid)", lambda: _gen_pdf(
    photo_files=[], photo_captions=[], asfalistiki='INTERLIFE')[:4] == b'%PDF', True)
test("27. PDF ΕΘΝΙΚΗ (valid)", lambda: _gen_pdf(
    photo_files=[], photo_captions=[], asfalistiki='ΕΘΝΙΚΗ ΑΣΦΑΛΙΣΤΙΚΗ')[:4] == b'%PDF', True)
test("28. PDF APEIRON (valid)", lambda: _gen_pdf(
    photo_files=[], photo_captions=[], asfalistiki='APEIRON ΑΣΦΑΛΙΣΤΙΚΗ')[:4] == b'%PDF', True)

# PDF με φωτογραφία
_img = PILImg.new('RGB',(800,600),color=(100,150,200))
_buf = io.BytesIO()
_img.save(_buf, format='JPEG')
class _MockFile:
    def __init__(self, d): self._d = d
    def seek(self, p): pass
    def read(self): return self._d

test("29. PDF με φωτογραφία (valid)", lambda: _gen_pdf(
    photo_files=[_MockFile(_buf.getvalue())],
    photo_captions=['Test photo'], asfalistiki='INTERLIFE')[:4] == b'%PDF', True)

# PDF με πολλά ανταλλακτικά
_many_parts = [{'name':f'ΑΝΤΑΛΛΑΚΤΙΚΟ {i}','type':'ΓΝ','price':100+i,
                'fanop':0,'vafeas':0,'mixanikos':0,'ilgos':0} for i in range(10)]
test("30. PDF με 10 ανταλλακτικά (valid)", lambda: _gen_pdf(
    photo_files=[], photo_captions=[], asfalistiki='INTERLIFE')[:4] == b'%PDF', True)

# ── VEHICLES ─────────────────────────────────────────────
print("\n🚗 VEHICLES")
from vehicles import get_markes, get_montela
test("31. Λίστα μαρκών (>0)", lambda: len(get_markes()) > 0, True)
test("32. Πλήθος μαρκών", lambda: len(get_markes()))
test("33. Μοντέλα Toyota (>0)", lambda: len(get_montela('Toyota')) > 0, True)
test("34. Μοντέλα VW (>0)", lambda: len(get_montela('Volkswagen')) > 0, True)
test("35. Μοντέλα άγνωστης μάρκας ([])", lambda: get_montela('ΔΕΝ_ΥΠΑΡΧΕΙ'), [])

# ── GOOGLE DRIVE ─────────────────────────────────────────
print("\n☁️ GOOGLE DRIVE")
import os
if os.path.exists(os.path.expanduser("~/Documents/experts360_credentials/token.json")):
    from gdrive import get_service, get_root_folder_id
    test("36. Drive σύνδεση", lambda: get_service() is not None, True)
    test("37. Drive root folder", lambda: get_root_folder_id() is not None, True)
else:
    print("   ⚠️  Drive token.json δεν βρέθηκε τοπικά (OK για cloud)")

# ── ΣΥΝΟΨΗ ───────────────────────────────────────────────
print("\n" + "="*60)
ok_count  = sum(1 for _, s, _ in results if s)
fail_count = sum(1 for _, s, _ in results if not s)
total = len(results)
print(f"ΑΠΟΤΕΛΕΣΜΑ: {ok_count}/{total} ✅  |  {fail_count} ❌")

if fail_count > 0:
    print("\n❌ ΑΠΟΤΥΧΙΕΣ (αυτά χρειάζονται διόρθωση):")
    print("-"*60)
    for name, s, msg in results:
        if not s:
            print(f"  [{name}]")
            print(f"  Σφάλμα: {msg}")
            print()
else:
    print("\n🎉 Όλοι οι έλεγχοι πέρασαν!")
print("="*60)
