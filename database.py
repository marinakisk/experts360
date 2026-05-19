"""
database.py - Διαχείριση βάσης δεδομένων για GNOMON EXPERTS
Υποστηρίζει SQLite (τοπικά) και PostgreSQL (Supabase/cloud)
"""

import sqlite3
import json
import os
from datetime import datetime
from typing import Optional, List, Dict, Any

# ============================================================
# ΡΥΘΜΙΣΕΙΣ
# ============================================================
SQLITE_PATH = "gnomon_db.sqlite"

# Για PostgreSQL: βάλε το connection string στο .env ή εδώ
# π.χ. "postgresql://user:pass@db.supabase.co:5432/postgres"
# Διαβάζει από Streamlit secrets ή environment variable
def _get_db_url():
    try:
        import streamlit as st
        return st.secrets.get("GNOMON_DB_URL", "") or os.environ.get("GNOMON_DB_URL", "")
    except:
        return os.environ.get("GNOMON_DB_URL", "")

POSTGRES_URL = _get_db_url()


def get_connection():
    """Επιστρέφει σύνδεση στη βάση (PostgreSQL αν υπάρχει URL, αλλιώς SQLite)."""
    if POSTGRES_URL:
        try:
            import psycopg2
            conn = psycopg2.connect(POSTGRES_URL)
            conn.autocommit = False
            return conn, "postgres"
        except Exception as e:
            print(f"PostgreSQL failed: {e}, falling back to SQLite")
    conn = sqlite3.connect(SQLITE_PATH)
    conn.row_factory = sqlite3.Row
    return conn, "sqlite"


def placeholder(db_type: str, n: int = 1) -> str:
    """Επιστρέφει το σωστό placeholder (%s για PG, ? για SQLite)."""
    if db_type == "postgres":
        return ", ".join(["%s"] * n)
    return ", ".join(["?"] * n)


def ph(db_type: str) -> str:
    return "%s" if db_type == "postgres" else "?"


# ============================================================
# ΔΗΜΙΟΥΡΓΙΑ ΠΙΝΑΚΩΝ
# ============================================================
SCHEMA_KTIRION_SQLITE = """
CREATE TABLE IF NOT EXISTS ektheseis_ktirion (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    arithmos_zimias TEXT,
    hm_epitheorisis TEXT,
    hm_syntaxis     TEXT,
    hm_anathesis    TEXT,
    hm_zimias       TEXT,
    pathon          TEXT,
    up_opsin        TEXT,
    asfalismenos    TEXT,
    asfalisterio    TEXT,
    topothesia      TEXT,
    thema           TEXT,
    eisagogi        TEXT,
    genikes         TEXT,
    eisagogi_diap   TEXT,
    bullets         TEXT,
    protasi         TEXT,
    prosopiki       TEXT,
    apaitisi_par    TEXT,
    total_apaitisi  REAL DEFAULT 0,
    total_ektimisi  REAL DEFAULT 0,
    asfalistiki     TEXT,
    onomateponymo   TEXT,
    status          TEXT DEFAULT 'draft',
    created_at      TEXT DEFAULT (datetime('now','localtime')),
    updated_at      TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS grammes_zimias_ktiriou (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ekthesi_id  INTEGER REFERENCES ektheseis_ktirion(id) ON DELETE CASCADE,
    desc        TEXT,
    apaitisi    REAL DEFAULT 0,
    ektimisi    REAL DEFAULT 0,
    sort_order  INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS history_ktirion (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ekthesi_id  INTEGER REFERENCES ektheseis_ktirion(id) ON DELETE CASCADE,
    action      TEXT,
    details     TEXT,
    user_name   TEXT,
    created_at  TEXT DEFAULT (datetime('now','localtime'))
);

CREATE INDEX IF NOT EXISTS idx_ktirion_arithmos ON ektheseis_ktirion(arithmos_zimias);
CREATE INDEX IF NOT EXISTS idx_ktirion_pathon   ON ektheseis_ktirion(pathon);
CREATE INDEX IF NOT EXISTS idx_ktirion_status   ON ektheseis_ktirion(status);
"""

SCHEMA_KTIRION_POSTGRES = """
CREATE TABLE IF NOT EXISTS ektheseis_ktirion (
    id              SERIAL PRIMARY KEY,
    arithmos_zimias TEXT,
    hm_epitheorisis TEXT,
    hm_syntaxis     TEXT,
    hm_anathesis    TEXT,
    hm_zimias       TEXT,
    pathon          TEXT,
    up_opsin        TEXT,
    asfalismenos    TEXT,
    asfalisterio    TEXT,
    topothesia      TEXT,
    thema           TEXT,
    eisagogi        TEXT,
    genikes         TEXT,
    eisagogi_diap   TEXT,
    bullets         TEXT,
    protasi         TEXT,
    prosopiki       TEXT,
    apaitisi_par    TEXT,
    total_apaitisi  REAL DEFAULT 0,
    total_ektimisi  REAL DEFAULT 0,
    asfalistiki     TEXT,
    onomateponymo   TEXT,
    status          TEXT DEFAULT 'draft',
    created_at      TIMESTAMP DEFAULT NOW(),
    updated_at      TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS grammes_zimias_ktiriou (
    id          SERIAL PRIMARY KEY,
    ekthesi_id  INTEGER REFERENCES ektheseis_ktirion(id) ON DELETE CASCADE,
    desc        TEXT,
    apaitisi    REAL DEFAULT 0,
    ektimisi    REAL DEFAULT 0,
    sort_order  INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS history_ktirion (
    id          SERIAL PRIMARY KEY,
    ekthesi_id  INTEGER REFERENCES ektheseis_ktirion(id) ON DELETE CASCADE,
    action      TEXT,
    details     TEXT,
    user_name   TEXT,
    created_at  TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_ktirion_arithmos ON ektheseis_ktirion(arithmos_zimias);
CREATE INDEX IF NOT EXISTS idx_ktirion_pathon   ON ektheseis_ktirion(pathon);
CREATE INDEX IF NOT EXISTS idx_ktirion_status   ON ektheseis_ktirion(status);
"""

SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS ektheseis (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ar_zimias   TEXT,
    hm_entolhs  TEXT,
    hm_atyx     TEXT,
    topos_atyx  TEXT,
    idioktitis  TEXT,
    ar_kykl     TEXT,
    ar_plaisiou TEXT,
    marka       TEXT,
    montelo     TEXT,
    kyvika      TEXT,
    xrisi       TEXT,
    proti_adeia TEXT,
    xiliom      TEXT,
    axia        INTEGER,
    hm_kteo     TEXT,
    visit_date  TEXT,
    visit_place TEXT,
    paratiriseis TEXT,
    onomateponymo TEXT,
    created_at  TEXT DEFAULT (datetime('now','localtime')),
    updated_at  TEXT DEFAULT (datetime('now','localtime')),
    status      TEXT DEFAULT 'draft'
);

CREATE TABLE IF NOT EXISTS grammes_antallaktikon (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ekthesi_id  INTEGER REFERENCES ektheseis(id) ON DELETE CASCADE,
    name        TEXT,
    type        TEXT,
    price       REAL DEFAULT 0,
    fanop       REAL DEFAULT 0,
    vafeas      REAL DEFAULT 0,
    mixanikos   REAL DEFAULT 0,
    ilgos       REAL DEFAULT 0,
    sort_order  INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS grammes_ergasion (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ekthesi_id  INTEGER REFERENCES ektheseis(id) ON DELETE CASCADE,
    type        TEXT,
    desc        TEXT,
    fanop       REAL DEFAULT 0,
    vafeas      REAL DEFAULT 0,
    mixanikos   REAL DEFAULT 0,
    ilgos       REAL DEFAULT 0,
    sort_order  INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS history (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ekthesi_id  INTEGER REFERENCES ektheseis(id) ON DELETE CASCADE,
    action      TEXT,
    details     TEXT,
    user_name   TEXT,
    created_at  TEXT DEFAULT (datetime('now','localtime'))
);

CREATE INDEX IF NOT EXISTS idx_ektheseis_ar_zimias ON ektheseis(ar_zimias);
CREATE INDEX IF NOT EXISTS idx_ektheseis_idioktitis ON ektheseis(idioktitis);
CREATE INDEX IF NOT EXISTS idx_ektheseis_marka ON ektheseis(marka);
CREATE INDEX IF NOT EXISTS idx_ektheseis_status ON ektheseis(status);
"""

SCHEMA_POSTGRES = """
CREATE TABLE IF NOT EXISTS ektheseis (
    id          SERIAL PRIMARY KEY,
    ar_zimias   TEXT,
    hm_entolhs  TEXT,
    hm_atyx     TEXT,
    topos_atyx  TEXT,
    idioktitis  TEXT,
    ar_kykl     TEXT,
    ar_plaisiou TEXT,
    marka       TEXT,
    montelo     TEXT,
    kyvika      TEXT,
    xrisi       TEXT,
    proti_adeia TEXT,
    xiliom      TEXT,
    axia        INTEGER,
    hm_kteo     TEXT,
    visit_date  TEXT,
    visit_place TEXT,
    paratiriseis TEXT,
    onomateponymo TEXT,
    created_at  TIMESTAMP DEFAULT NOW(),
    updated_at  TIMESTAMP DEFAULT NOW(),
    status      TEXT DEFAULT 'draft'
);

CREATE TABLE IF NOT EXISTS grammes_antallaktikon (
    id          SERIAL PRIMARY KEY,
    ekthesi_id  INTEGER REFERENCES ektheseis(id) ON DELETE CASCADE,
    name        TEXT,
    type        TEXT,
    price       REAL DEFAULT 0,
    fanop       REAL DEFAULT 0,
    vafeas      REAL DEFAULT 0,
    mixanikos   REAL DEFAULT 0,
    ilgos       REAL DEFAULT 0,
    sort_order  INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS grammes_ergasion (
    id          SERIAL PRIMARY KEY,
    ekthesi_id  INTEGER REFERENCES ektheseis(id) ON DELETE CASCADE,
    type        TEXT,
    desc        TEXT,
    fanop       REAL DEFAULT 0,
    vafeas      REAL DEFAULT 0,
    mixanikos   REAL DEFAULT 0,
    ilgos       REAL DEFAULT 0,
    sort_order  INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS history (
    id          SERIAL PRIMARY KEY,
    ekthesi_id  INTEGER REFERENCES ektheseis(id) ON DELETE CASCADE,
    action      TEXT,
    details     TEXT,
    user_name   TEXT,
    created_at  TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_ektheseis_ar_zimias  ON ektheseis(ar_zimias);
CREATE INDEX IF NOT EXISTS idx_ektheseis_idioktitis ON ektheseis(idioktitis);
CREATE INDEX IF NOT EXISTS idx_ektheseis_marka      ON ektheseis(marka);
CREATE INDEX IF NOT EXISTS idx_ektheseis_status     ON ektheseis(status);
"""


def init_db():
    """Δημιουργεί τους πίνακες αν δεν υπάρχουν."""
    conn, db_type = get_connection()
    try:
        cur = conn.cursor()
        if db_type == "postgres":
            for stmt in (SCHEMA_POSTGRES + SCHEMA_KTIRION_POSTGRES).split(";"):
                s = stmt.strip()
                if s:
                    cur.execute(s)
        else:
            conn.executescript(SCHEMA_SQLITE + SCHEMA_KTIRION_SQLITE)
        conn.commit()
        return True, db_type
    except Exception as e:
        return False, str(e)
    finally:
        conn.close()


# ============================================================
# ΑΠΟΘΗΚΕΥΣΗ / ΕΝΗΜΕΡΩΣΗ
# ============================================================
def save_ekthesi(data: dict, parts: list, works: list,
                 user_name: str = "", ekthesi_id: int = None) -> tuple:
    """
    Αποθηκεύει ή ενημερώνει έκθεση.
    Επιστρέφει (id, error_message).
    """
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()

        fields = {
            "ar_zimias":   data.get("ar_zimias",""),
            "hm_entolhs":  data.get("hm_entolhs",""),
            "hm_atyx":     data.get("hm_atyxhmatos",""),
            "topos_atyx":  data.get("topos_atyxhmatos",""),
            "idioktitis":  data.get("idioktitis",""),
            "ar_kykl":     data.get("ar_kykloforias",""),
            "marka":       data.get("marka",""),
            "montelo":     data.get("montelo",""),
            "kyvika":      data.get("kyvika",""),
            "xrisi":       data.get("xrisi",""),
            "proti_adeia": data.get("proti_adeia",""),
            "xiliom":      data.get("xiliometrites",""),
            "axia":        int(data.get("axia",0) or 0),
            "hm_kteo":     data.get("hm_kteo",""),
            "visit_date":  data.get("visit_date",""),
            "visit_place": data.get("visit_place",""),
            "paratiriseis":data.get("paratiriseis",""),
            "onomateponymo":data.get("onomateponymo",""),
            "status":      data.get("status","draft"),
        }

        now = datetime.now().strftime("%d/%m/%Y %H:%M")

        if ekthesi_id:
            # UPDATE
            set_clause = ", ".join([f"{k}={p}" for k in fields])
            set_clause += f", updated_at={p}"
            vals = list(fields.values()) + [now, ekthesi_id]
            cur.execute(f"UPDATE ektheseis SET {set_clause} WHERE id={p}", vals)
            action = "ΕΝΗΜΕΡΩΣΗ"
        else:
            # INSERT
            cols = ", ".join(fields.keys())
            phs  = ", ".join([p]*len(fields))
            cur.execute(f"INSERT INTO ektheseis ({cols}) VALUES ({phs})", list(fields.values()))
            if db_type == "postgres":
                cur.execute("SELECT lastval()")
                ekthesi_id = cur.fetchone()[0]
            else:
                ekthesi_id = cur.lastrowid
            action = "ΔΗΜΙΟΥΡΓΙΑ"

        # Διαγραφή παλιών γραμμών
        cur.execute(f"DELETE FROM grammes_antallaktikon WHERE ekthesi_id={p}", [ekthesi_id])
        cur.execute(f"DELETE FROM grammes_ergasion WHERE ekthesi_id={p}", [ekthesi_id])

        # Εισαγωγή ανταλλακτικών
        for i, part in enumerate(parts):
            if not part.get("name","").strip():
                continue
            cur.execute(f"""
                INSERT INTO grammes_antallaktikon
                (ekthesi_id,name,type,price,fanop,vafeas,mixanikos,ilgos,sort_order)
                VALUES ({placeholder(db_type,9)})
            """, [ekthesi_id, part.get("name",""), part.get("type",""),
                  part.get("price",0), part.get("fanop",0), part.get("vafeas",0),
                  part.get("mixanikos",0), part.get("ilgos",0), i])

        # Εισαγωγή εργασιών
        for i, work in enumerate(works):
            if not work.get("type","").strip() and not work.get("desc","").strip():
                continue
            cur.execute(f"""
                INSERT INTO grammes_ergasion
                (ekthesi_id,type,desc,fanop,vafeas,mixanikos,ilgos,sort_order)
                VALUES ({placeholder(db_type,8)})
            """, [ekthesi_id, work.get("type",""), work.get("desc",""),
                  work.get("fanop",0), work.get("vafeas",0),
                  work.get("mixanikos",0), work.get("ilgos",0), i])

        # Ιστορικό
        details = f"Αρ.Ζημίας: {fields['ar_zimias']} | {fields['idioktitis']} | {fields['marka']} {fields['montelo']}"
        cur.execute(f"""
            INSERT INTO history (ekthesi_id, action, details, user_name)
            VALUES ({placeholder(db_type,4)})
        """, [ekthesi_id, action, details, user_name])

        conn.commit()
        return ekthesi_id, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


# ============================================================
# ΑΝΑΖΗΤΗΣΗ
# ============================================================
def search_ektheseis(query: str = "", status: str = "",
                     marka: str = "", montelo: str = "",
                     ar_kykl: str = "", asfalistiki: str = "",
                     hm_from: str = "", hm_to: str = "",
                     limit: int = 100) -> List[Dict]:
    """Αναζήτηση εκθέσεων με φίλτρα."""
    conn, db_type = get_connection()
    p = ph(db_type)
    like = "ILIKE" if db_type == "postgres" else "LIKE"
    try:
        conditions, params = [], []

        if query:
            conditions.append(f"(ar_zimias {like} {p} OR idioktitis {like} {p} OR ar_kykl {like} {p})")
            q = f"%{query}%"
            params += [q, q, q]
        if status:
            conditions.append(f"status = {p}"); params.append(status)
        if marka:
            conditions.append(f"marka = {p}"); params.append(marka)
        if montelo:
            conditions.append(f"montelo = {p}"); params.append(montelo)
        if ar_kykl:
            conditions.append(f"ar_kykl {like} {p}"); params.append(f"%{ar_kykl}%")
        if asfalistiki:
            conditions.append(f"asfalistiki = {p}"); params.append(asfalistiki)
        if hm_from:
            conditions.append(f"hm_entolhs >= {p}"); params.append(hm_from)
        if hm_to:
            conditions.append(f"hm_entolhs <= {p}"); params.append(hm_to)

        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""

        cur = conn.cursor()
        cur.execute(f"""
            SELECT id, ar_zimias, hm_entolhs, hm_atyx, idioktitis,
                   ar_kykl, marka, montelo, axia, status,
                   asfalistiki, created_at
            FROM ektheseis {where}
            ORDER BY id DESC LIMIT {limit}
        """, params)
        return [dict(r) for r in cur.fetchall()]
    except Exception as e:
        print(f"Search error: {e}")
        return []
    finally:
        conn.close()


# ============================================================
# ΦΟΡΤΩΣΗ ΕΚΘΕΣΗΣ
# ============================================================
def load_ekthesi(ekthesi_id: int) -> Optional[Dict]:
    """Φορτώνει πλήρη έκθεση με ανταλλακτικά και εργασίες."""
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()

        cur.execute(f"SELECT * FROM ektheseis WHERE id={p}", [ekthesi_id])
        row = cur.fetchone()
        if not row:
            return None
        data = dict(row)

        cur.execute(f"""
            SELECT * FROM grammes_antallaktikon
            WHERE ekthesi_id={p} ORDER BY sort_order
        """, [ekthesi_id])
        data["parts"] = [dict(r) for r in cur.fetchall()]

        cur.execute(f"""
            SELECT * FROM grammes_ergasion
            WHERE ekthesi_id={p} ORDER BY sort_order
        """, [ekthesi_id])
        data["works"] = [dict(r) for r in cur.fetchall()]

        cur.execute(f"""
            SELECT * FROM history WHERE ekthesi_id={p} ORDER BY id DESC LIMIT 20
        """, [ekthesi_id])
        data["history"] = [dict(r) for r in cur.fetchall()]

        return data
    except Exception as e:
        print(f"Load error: {e}")
        return None
    finally:
        conn.close()


# ============================================================
# ΔΙΑΓΡΑΦΗ
# ============================================================
def delete_ekthesi(ekthesi_id: int, user_name: str = "") -> tuple:
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        cur.execute(f"SELECT ar_zimias, idioktitis FROM ektheseis WHERE id={p}", [ekthesi_id])
        row = cur.fetchone()
        if not row:
            return False, "Δεν βρέθηκε η έκθεση"
        cur.execute(f"DELETE FROM ektheseis WHERE id={p}", [ekthesi_id])
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


# ============================================================
# ΣΤΑΤΙΣΤΙΚΑ
# ============================================================
def get_statistics() -> Dict:
    conn, db_type = get_connection()
    try:
        cur = conn.cursor()
        stats = {}

        cur.execute("SELECT COUNT(*) as total FROM ektheseis")
        stats["total"] = cur.fetchone()[0]

        cur.execute("""
            SELECT COUNT(*) FROM ektheseis
            WHERE created_at >= date('now', '-30 days')
        """ if db_type == "sqlite" else """
            SELECT COUNT(*) FROM ektheseis
            WHERE created_at >= NOW() - INTERVAL '30 days'
        """)
        stats["last_30_days"] = cur.fetchone()[0]

        cur.execute("SELECT COALESCE(SUM(axia),0) FROM ektheseis")
        stats["total_axia"] = cur.fetchone()[0]

        cur.execute("""
            SELECT marka, COUNT(*) as cnt
            FROM ektheseis WHERE marka != ''
            GROUP BY marka ORDER BY cnt DESC LIMIT 5
        """)
        stats["top_markes"] = [dict(r) for r in cur.fetchall()]

        cur.execute("""
            SELECT status, COUNT(*) as cnt
            FROM ektheseis GROUP BY status
        """)
        stats["by_status"] = {r[0]: r[1] for r in cur.fetchall()}

        cur.execute("""
            SELECT COALESCE(SUM(e.axia),0) as total,
                   COUNT(*) as cnt,
                   e.hm_entolhs
            FROM ektheseis e
            WHERE e.hm_entolhs != ''
            GROUP BY substr(e.hm_entolhs,4,7)
            ORDER BY e.hm_entolhs DESC LIMIT 6
        """)
        stats["by_month"] = [dict(r) for r in cur.fetchall()]

        return stats
    except Exception as e:
        print(f"Stats error: {e}")
        return {}
    finally:
        conn.close()


# ============================================================
# ΕΝΗΜΕΡΩΣΗ STATUS
# ============================================================
def update_status(ekthesi_id: int, status: str, user_name: str = "") -> tuple:
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        cur.execute(f"UPDATE ektheseis SET status={p}, updated_at={p} WHERE id={p}",
                    [status, now, ekthesi_id])
        cur.execute(f"""
            INSERT INTO history (ekthesi_id, action, details, user_name)
            VALUES ({placeholder(db_type,4)})
        """, [ekthesi_id, "ΑΛΛΑΓΗ STATUS", status, user_name])
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


# ============================================================
# ΙΣΤΟΡΙΚΟ
# ============================================================
def get_history(ekthesi_id: int) -> List[Dict]:
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        cur.execute(f"""
            SELECT * FROM history WHERE ekthesi_id={p}
            ORDER BY id DESC
        """, [ekthesi_id])
        return [dict(r) for r in cur.fetchall()]
    except:
        return []
    finally:
        conn.close()


# ============================================================
# ΕΞΑΓΩΓΗ ΟΛΩΝ ΣΕ JSON (backup)
# ============================================================
def export_all_json() -> str:
    conn, db_type = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM ektheseis ORDER BY id")
        ektheseis = [dict(r) for r in cur.fetchall()]
        for e in ektheseis:
            cur.execute("SELECT * FROM grammes_antallaktikon WHERE ekthesi_id=?", [e["id"]])
            e["parts"] = [dict(r) for r in cur.fetchall()]
            cur.execute("SELECT * FROM grammes_ergasion WHERE ekthesi_id=?", [e["id"]])
            e["works"] = [dict(r) for r in cur.fetchall()]
        return json.dumps(ektheseis, ensure_ascii=False, indent=2)
    finally:
        conn.close()


# ============================================================
# ΕΜΠΟΡΙΚΗ ΑΞΙΑ - ΜΝΗΜΗ ΑΠΟ ΠΡΟΗΓΟΥΜΕΝΕΣ ΕΚΘΕΣΕΙΣ
# ============================================================
def get_axia_stats(marka: str, montelo: str) -> dict:
    if not marka or not montelo:
        return {}
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        cur.execute(f"""
            SELECT axia, hm_entolhs FROM ektheseis
            WHERE marka={p} AND montelo={p} AND axia > 0
            ORDER BY id DESC LIMIT 20
        """, [marka, montelo])
        rows = cur.fetchall()
        if not rows:
            return {}
        axies = [r[0] for r in rows]
        return {
            'count': len(axies),
            'mean':  int(sum(axies)/len(axies)),
            'min':   min(axies),
            'max':   max(axies),
            'last':  axies[0],
            'last_date': rows[0][1] if rows[0][1] else '',
        }
    except:
        return {}
    finally:
        conn.close()


def get_antallaktiko_stats(marka: str, montelo: str, antallaktiko: str,
                           typos: str = "") -> dict:
    """
    Επιστρέφει στατιστικά τιμής ανταλλακτικού για μάρκα/μοντέλο.
    Φιλτράρει υποχρεωτικά ανά τύπο (ΓΝ / ΙΜ / ΜΤΧ) αν δοθεί.
    """
    if not marka or not montelo or not antallaktiko or len(antallaktiko) < 3:
        return {}
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        like_op = "ILIKE" if db_type == "postgres" else "LIKE"

        params = [marka, montelo, f"%{antallaktiko}%"]
        type_clause = ""
        if typos:
            type_clause = f"AND ga.type = {p}"
            params.append(typos)

        cur.execute(f"""
            SELECT ga.price, ga.type FROM grammes_antallaktikon ga
            JOIN ektheseis e ON e.id = ga.ekthesi_id
            WHERE e.marka={p} AND e.montelo={p}
            AND ga.name {like_op} {p}
            AND ga.price > 0
            {type_clause}
            ORDER BY ga.id DESC LIMIT 20
        """, params)

        rows = cur.fetchall()
        if not rows:
            return {}

        prices = [r[0] for r in rows]
        return {
            'count': len(prices),
            'mean':  round(sum(prices)/len(prices), 2),
            'min':   min(prices),
            'max':   max(prices),
            'typos': typos if typos else "Όλοι τύποι",
        }
    except:
        return {}
    finally:
        conn.close()


def get_antallaktiko_stats_all_types(marka: str, montelo: str,
                                     antallaktiko: str) -> dict:
    """Επιστρέφει στατιστικά χωριστά ανά τύπο (ΓΝ/ΙΜ/ΜΤΧ)."""
    result = {}
    for t in ["ΓΝ", "ΙΜ", "ΜΤΧ"]:
        stats = get_antallaktiko_stats(marka, montelo, antallaktiko, t)
        if stats:
            result[t] = stats
    return result


# ============================================================
# ΚΤΙΡΙΑ - ΑΠΟΘΗΚΕΥΣΗ / ΦΟΡΤΩΣΗ / ΑΝΑΖΗΤΗΣΗ
# ============================================================
def save_ekthesi_ktiriou(data: dict, grammes: list,
                          user_name: str = "", ekthesi_id: int = None) -> tuple:
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        fields = {
            "arithmos_zimias": data.get("arithmos_zimias", ""),
            "hm_epitheorisis": data.get("hm_epitheorisis", ""),
            "hm_syntaxis":     data.get("hm_syntaxis", ""),
            "hm_anathesis":    data.get("hm_anathesis", ""),
            "hm_zimias":       data.get("hm_zimias", ""),
            "pathon":          data.get("pathon", ""),
            "up_opsin":        data.get("up_opsin", ""),
            "asfalismenos":    data.get("asfalismenos", ""),
            "asfalisterio":    data.get("asfalisterio", ""),
            "topothesia":      data.get("topothesia", ""),
            "thema":           data.get("thema", ""),
            "eisagogi":        data.get("eisagogi", ""),
            "genikes":         data.get("genikes", ""),
            "eisagogi_diap":   data.get("eisagogi_diap", ""),
            "bullets":         data.get("bullets", ""),
            "protasi":         data.get("protasi", ""),
            "prosopiki":       data.get("prosopiki", ""),
            "apaitisi_par":    data.get("apaitisi_par", ""),
            "total_apaitisi":  float(data.get("total_apaitisi", 0)),
            "total_ektimisi":  float(data.get("total_ektimisi", 0)),
            "asfalistiki":     data.get("asfalistiki", ""),
            "onomateponymo":   data.get("onomateponymo", ""),
            "status":          data.get("status", "draft"),
        }
        now = datetime.now().strftime("%d/%m/%Y %H:%M")

        if ekthesi_id:
            set_clause = ", ".join([f"{k}={p}" for k in fields])
            set_clause += f", updated_at={p}"
            vals = list(fields.values()) + [now, ekthesi_id]
            cur.execute(f"UPDATE ektheseis_ktirion SET {set_clause} WHERE id={p}", vals)
            action = "ΕΝΗΜΕΡΩΣΗ"
        else:
            cols = ", ".join(fields.keys())
            phs  = ", ".join([p]*len(fields))
            cur.execute(f"INSERT INTO ektheseis_ktirion ({cols}) VALUES ({phs})",
                        list(fields.values()))
            if db_type == "postgres":
                cur.execute("SELECT lastval()")
            ekthesi_id = cur.lastrowid if db_type == "sqlite" else cur.fetchone()[0]
            action = "ΔΗΜΙΟΥΡΓΙΑ"

        cur.execute(f"DELETE FROM grammes_zimias_ktiriou WHERE ekthesi_id={p}", [ekthesi_id])
        for i, g in enumerate(grammes):
            if not g.get("desc", "").strip():
                continue
            cur.execute(f"""
                INSERT INTO grammes_zimias_ktiriou
                (ekthesi_id, desc, apaitisi, ektimisi, sort_order)
                VALUES ({placeholder(db_type, 5)})
            """, [ekthesi_id, g.get("desc",""), float(g.get("apaitisi",0)),
                  float(g.get("ektimisi",0)), i])

        details = f"Αρ.Ζημίας: {fields['arithmos_zimias']} | {fields['pathon']} | {fields['topothesia'][:30]}"
        cur.execute(f"""
            INSERT INTO history_ktirion (ekthesi_id, action, details, user_name)
            VALUES ({placeholder(db_type,4)})
        """, [ekthesi_id, action, details, user_name])

        conn.commit()
        return ekthesi_id, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


def load_ekthesi_ktiriou(ekthesi_id: int) -> dict:
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM ektheseis_ktirion WHERE id={p}", [ekthesi_id])
        row = cur.fetchone()
        if not row:
            return None
        data = dict(row)
        cur.execute(f"""
            SELECT * FROM grammes_zimias_ktiriou
            WHERE ekthesi_id={p} ORDER BY sort_order
        """, [ekthesi_id])
        data["grammes"] = [dict(r) for r in cur.fetchall()]
        cur.execute(f"""
            SELECT * FROM history_ktirion WHERE ekthesi_id={p} ORDER BY id DESC LIMIT 10
        """, [ekthesi_id])
        data["history"] = [dict(r) for r in cur.fetchall()]
        return data
    except:
        return None
    finally:
        conn.close()


def search_ektheseis_ktirion(query: str = "", status: str = "", limit: int = 50) -> list:
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        conditions, params = [], []
        if query:
            like_op = "ILIKE" if db_type == "postgres" else "LIKE"
            conditions.append(f"""(arithmos_zimias {like_op} {p} OR
                                   pathon {like_op} {p} OR
                                   topothesia {like_op} {p})""")
            q = f"%{query}%"
            params += [q, q, q]
        if status:
            conditions.append(f"status = {p}")
            params.append(status)
        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        params.append(limit)
        cur = conn.cursor()
        cur.execute(f"""
            SELECT id, arithmos_zimias, hm_epitheorisis, pathon, topothesia,
                   total_apaitisi, total_ektimisi, asfalistiki, status, created_at
            FROM ektheseis_ktirion
            {where} ORDER BY id DESC LIMIT {p}
        """, params)
        return [dict(r) for r in cur.fetchall()]
    except Exception as e:
        print(f"Search ktirion error: {e}")
        return []
    finally:
        conn.close()


def delete_ekthesi_ktiriou(ekthesi_id: int) -> tuple:
    conn, db_type = get_connection()
    p = ph(db_type)
    try:
        cur = conn.cursor()
        cur.execute(f"DELETE FROM ektheseis_ktirion WHERE id={p}", [ekthesi_id])
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def get_statistics_ktirion() -> dict:
    conn, db_type = get_connection()
    try:
        cur = conn.cursor()
        stats = {}
        cur.execute("SELECT COUNT(*) FROM ektheseis_ktirion")
        stats["total"] = cur.fetchone()[0]
        cur.execute("SELECT COALESCE(SUM(total_ektimisi),0) FROM ektheseis_ktirion")
        stats["total_ektimisi"] = cur.fetchone()[0]
        cur.execute("""
            SELECT COUNT(*) FROM ektheseis_ktirion
            WHERE created_at >= date('now','-30 days')
        """ if db_type == "sqlite" else """
            SELECT COUNT(*) FROM ektheseis_ktirion
            WHERE created_at >= NOW() - INTERVAL '30 days'
        """)
        stats["last_30_days"] = cur.fetchone()[0]
        cur.execute("""
            SELECT asfalistiki, COUNT(*) as cnt
            FROM ektheseis_ktirion WHERE asfalistiki != ''
            GROUP BY asfalistiki ORDER BY cnt DESC
        """)
        stats["by_asfalistiki"] = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT status, COUNT(*) FROM ektheseis_ktirion GROUP BY status")
        stats["by_status"] = {r[0]: r[1] for r in cur.fetchall()}
        return stats
    except Exception as e:
        print(f"Stats ktirion error: {e}")
        return {}
    finally:
        conn.close()
