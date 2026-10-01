# -*- coding: utf-8 -*-
"""
bh_db - SQLite baza: sxema, migratsiya, asosiy so'rovlar.

Baza HAQIQIY MANBA. Excel fayllar faqat kirish/chiqish.
Shu tufayli:
  - yillararo qoldiq o'z-o'zidan o'tadi
  - bir xil faylni qayta tashlasangiz dublikat bo'lmaydi
  - buxgalter bir marta qo'lda biriktirgan juftlik esda qoladi
"""

import os
import json
import sqlite3
import hashlib
import datetime
from decimal import Decimal

import bh_config as C


# ---------------------------------------------------------------------------
# Decimal <-> SQLite
# SQLite REAL ishlatsak tiyinlar siljiydi. Shuning uchun pul TEXT sifatida
# saqlanadi va Decimal bo'lib qaytariladi.
# ---------------------------------------------------------------------------
sqlite3.register_adapter(Decimal, lambda d: str(d))


def D(v, default="0"):
    if v is None or v == "":
        return Decimal(default)
    if isinstance(v, Decimal):
        return v
    return Decimal(str(v))


MIGRATIONS = {}


def migration(version):
    def deco(fn):
        MIGRATIONS[version] = fn
        return fn
    return deco


def run_script(cx, sql):
    """
    SQL matnini gap-gap bajaradi.

    `executescript` ishlatilmaydi: u ochiq tranzaksiyani avtomatik COMMIT
    qilib yuboradi, natijada migratsiya atomik bo'lmay qoladi.
    """
    for stmt in sql.split(";"):
        s = stmt.strip()
        if s:
            cx.execute(s)


@migration(1)
def _m1(cx):
    run_script(cx, """
    CREATE TABLE IF NOT EXISTS meta (
        key   TEXT PRIMARY KEY,
        value TEXT
    );

    CREATE TABLE IF NOT EXISTS setting (
        key   TEXT PRIMARY KEY,
        value TEXT
    );

    -- Import qilingan fayllar. sha256 -> bir fayl ikki marta kirmaydi.
    CREATE TABLE IF NOT EXISTS source_file (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        sha256       TEXT NOT NULL UNIQUE,
        filename     TEXT NOT NULL,
        full_path    TEXT,
        kind         TEXT NOT NULL,          -- kirim | chiqim | legacy
        size_bytes   INTEGER,
        imported_at  TEXT NOT NULL,
        doc_count    INTEGER DEFAULT 0,
        line_count   INTEGER DEFAULT 0,
        note         TEXT
    );

    -- Hujjat: faktura yoki kassa cheki
    CREATE TABLE IF NOT EXISTS document (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        source_file_id INTEGER REFERENCES source_file(id) ON DELETE CASCADE,
        kind           TEXT NOT NULL,        -- kirim | chiqim
        doc_key        TEXT NOT NULL UNIQUE, -- mantiqiy kalit (dublikat qalqoni)
        doc_no         TEXT,
        doc_date       TEXT,                 -- ISO YYYY-MM-DD
        doc_year       INTEGER,
        contract_no    TEXT,
        contract_date  TEXT,
        partner_name   TEXT,
        partner_tin    TEXT,
        pos_id         TEXT,                 -- ФМ рақами (chiqim uchun)
        check_type     TEXT,                 -- Сотув | Қайтариш
        total_net      TEXT,
        total_vat      TEXT,
        total_gross    TEXT,
        is_return      INTEGER DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS ix_doc_year ON document(doc_year, kind);
    CREATE INDEX IF NOT EXISTS ix_doc_date ON document(doc_date);

    -- Hujjat satri
    CREATE TABLE IF NOT EXISTS doc_line (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id    INTEGER NOT NULL REFERENCES document(id) ON DELETE CASCADE,
        kind           TEXT NOT NULL,        -- kirim | chiqim
        line_no        INTEGER,
        raw_name       TEXT,
        norm_name      TEXT,
        mxik           TEXT,
        mxik_name      TEXT,
        barcode        TEXT,
        marking_code   TEXT,
        is_marked      INTEGER,
        unit_raw       TEXT,
        unit           TEXT,
        packaging      TEXT,
        qty            TEXT,
        unit_price_net TEXT,   -- QQS'siz birlik narxi (kirim)
        amount_net     TEXT,
        vat_amount     TEXT,
        amount_gross   TEXT,
        unit_price_gross TEXT, -- QQS bilan birlik narxi (chiqim)
        product_id     INTEGER REFERENCES product(id),
        match_method   TEXT,   -- barcode|marking|name|prefix|fuzzy|mxik|manual
        match_score    REAL,
        line_date      TEXT
    );
    CREATE INDEX IF NOT EXISTS ix_line_doc  ON doc_line(document_id);
    CREATE INDEX IF NOT EXISTS ix_line_prod ON doc_line(product_id);
    CREATE INDEX IF NOT EXISTS ix_line_kind ON doc_line(kind, product_id);
    CREATE INDEX IF NOT EXISTS ix_line_norm ON doc_line(norm_name);

    -- Mahsulot kartochkasi (kirim va chiqim shu yerda birlashadi)
    CREATE TABLE IF NOT EXISTS product (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        canon_name  TEXT NOT NULL,
        norm_name   TEXT NOT NULL,
        mxik        TEXT,
        mxik_name   TEXT,
        unit        TEXT,
        barcode     TEXT,
        is_marked   INTEGER,
        markup      TEXT,          -- shu mahsulot uchun ustama (bo'sh = yillik)
        created_at  TEXT,
        note        TEXT
    );
    CREATE INDEX IF NOT EXISTS ix_prod_norm ON product(norm_name);
    CREATE INDEX IF NOT EXISTS ix_prod_mxik ON product(mxik);
    CREATE INDEX IF NOT EXISTS ix_prod_bar  ON product(barcode);

    -- O'RGANILGAN moslik. Buxgalter bir marta tasdiqlaydi - qayta so'ralmaydi.
    CREATE TABLE IF NOT EXISTS alias (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER NOT NULL REFERENCES product(id) ON DELETE CASCADE,
        alias_norm TEXT NOT NULL UNIQUE,
        alias_raw  TEXT,
        source     TEXT,           -- auto | manual
        created_at TEXT
    );

    -- FIFO partiyalari (har kirim satri = bitta partiya)
    CREATE TABLE IF NOT EXISTS stock_lot (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id  INTEGER NOT NULL REFERENCES product(id) ON DELETE CASCADE,
        doc_line_id INTEGER REFERENCES doc_line(id) ON DELETE CASCADE,
        lot_date    TEXT NOT NULL,
        lot_year    INTEGER,
        qty_in      TEXT NOT NULL,
        qty_left    TEXT NOT NULL,
        unit_cost   TEXT NOT NULL,   -- QQS'siz tannarx
        markup      TEXT,
        is_opening  INTEGER DEFAULT 0,   -- o'tgan yildan ko'chirilgan qoldiqmi
        source_year INTEGER               -- qaysi yildan kelgan
    );
    CREATE INDEX IF NOT EXISTS ix_lot_prod ON stock_lot(product_id, lot_date, id);
    CREATE INDEX IF NOT EXISTS ix_lot_year ON stock_lot(lot_year);

    -- Sotuvning qaysi partiyadan yechilgani
    CREATE TABLE IF NOT EXISTS allocation (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_line_id INTEGER NOT NULL REFERENCES doc_line(id) ON DELETE CASCADE,
        stock_lot_id INTEGER REFERENCES stock_lot(id) ON DELETE CASCADE,
        product_id   INTEGER,
        alloc_date   TEXT,
        alloc_year   INTEGER,
        qty          TEXT NOT NULL,
        unit_cost    TEXT NOT NULL,
        cost_amount  TEXT NOT NULL,
        shortfall    INTEGER DEFAULT 0   -- ombor yetmadi, tannarx taxminiy
    );
    CREATE INDEX IF NOT EXISTS ix_alloc_sale ON allocation(sale_line_id);
    CREATE INDEX IF NOT EXISTS ix_alloc_year ON allocation(alloc_year, product_id);

    -- Tekshiruv natijalari
    CREATE TABLE IF NOT EXISTS issue (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        year       INTEGER,
        severity   TEXT NOT NULL,
        code       TEXT NOT NULL,
        message    TEXT NOT NULL,
        ref_table  TEXT,
        ref_id     INTEGER,
        detail     TEXT,
        created_at TEXT,
        resolved   INTEGER DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS ix_issue_year ON issue(year, resolved, severity);

    -- Yil bo'yicha ustama stavkasi
    CREATE TABLE IF NOT EXISTS markup_rule (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        year       INTEGER,
        mxik       TEXT,
        partner_tin TEXT,
        markup     TEXT NOT NULL,
        note       TEXT
    );
    """)


@migration(2)
def _m2(cx):
    """
    Har STIR alohida (1.5.0). Ilgari ikki do'kon fayllari bitta bazaga qo'shilsa,
    hisobot ikkalasini aralashtirib yuborardi (kirim = ikki do'kon yig'indisi) va
    FIFO bir do'kon sotuvini boshqasining kirimidan yechardi.

    document.owner_tin - hujjat kimniki: fakturada sotib oluvchi, chekda sotuvchi.
    Cheklar uchun darrov to'ldiriladi (partner_tin); fakturalar uchun - fayl qayta
    o'qilib (bh_fifo.backfill_owners), chunki sotib oluvchi STIRi ilgari saqlanmagan.
    """
    run_script(cx, """
    ALTER TABLE document ADD COLUMN owner_tin TEXT;
    CREATE INDEX IF NOT EXISTS ix_doc_owner ON document(owner_tin, kind, doc_year);
    CREATE TABLE IF NOT EXISTS org (
        tin   TEXT PRIMARY KEY,
        name  TEXT NOT NULL DEFAULT ''
    );
    UPDATE document SET owner_tin = partner_tin
     WHERE kind = 'chiqim' AND partner_tin IS NOT NULL AND partner_tin <> ''
    """)


SCHEMA_LATEST = max(MIGRATIONS)


# ---------------------------------------------------------------------------
# Ulanish
# ---------------------------------------------------------------------------
def connect(path=None):
    path = path or C.db_path()
    first = not os.path.exists(path)
    cx = sqlite3.connect(path, timeout=30, isolation_level=None)
    cx.row_factory = sqlite3.Row
    cx.execute("PRAGMA journal_mode=WAL")
    cx.execute("PRAGMA foreign_keys=ON")
    cx.execute("PRAGMA synchronous=NORMAL")
    migrate(cx, path, first)
    return cx


def _current_version(cx):
    try:
        cx.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
        r = cx.execute("SELECT value FROM meta WHERE key='schema_version'").fetchone()
        return int(r["value"]) if r else 0
    except Exception:
        return 0


def migrate(cx, path, first=False):
    """
    Sxemani oxirgi versiyaga keltiradi.
    HAR migratsiyadan oldin baza nusxasi olinadi - bitta noto'g'ri
    `git push` buxgalterning yillik ma'lumotini yo'q qilmasligi uchun.
    """
    cur = _current_version(cx)
    if cur >= SCHEMA_LATEST:
        return cur

    if not first and cur > 0:
        backup(path, suffix="pre-v%d" % SCHEMA_LATEST)

    for v in sorted(MIGRATIONS):
        if v > cur:
            cx.execute("BEGIN")
            try:
                MIGRATIONS[v](cx)
                cx.execute(
                    "INSERT INTO meta(key,value) VALUES('schema_version',?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    (str(v),),
                )
                cx.execute("COMMIT")
            except Exception:
                cx.execute("ROLLBACK")
                raise
    cx.execute(
        "INSERT INTO meta(key,value) VALUES('last_open',?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (datetime.datetime.now().isoformat(timespec="seconds"),),
    )
    return SCHEMA_LATEST


def backup(path, suffix=None):
    if not os.path.exists(path):
        return None
    import shutil

    stamp = suffix or datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    d = os.path.join(os.path.dirname(path), "backup")
    os.makedirs(d, exist_ok=True)
    dst = os.path.join(d, "hisobot-%s.db" % stamp)
    try:
        shutil.copy2(path, dst)
        return dst
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Sozlamalar
# ---------------------------------------------------------------------------
def get_setting(cx, key, default=None):
    r = cx.execute("SELECT value FROM setting WHERE key=?", (key,)).fetchone()
    if not r:
        return default
    try:
        return json.loads(r["value"])
    except Exception:
        return r["value"]


def set_setting(cx, key, value):
    cx.execute(
        "INSERT INTO setting(key,value) VALUES(?,?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, json.dumps(value, ensure_ascii=False)),
    )


def get_markup(cx, year, mxik=None, partner_tin=None):
    """Ustama stavkasi: mahsulot > ta'minotchi > yil > umumiy."""
    if mxik:
        r = cx.execute(
            "SELECT markup FROM markup_rule WHERE mxik=? AND (year IS NULL OR year=?) "
            "ORDER BY year DESC LIMIT 1", (mxik, year)).fetchone()
        if r:
            return D(r["markup"])
    if partner_tin:
        r = cx.execute(
            "SELECT markup FROM markup_rule WHERE partner_tin=? AND (year IS NULL OR year=?) "
            "ORDER BY year DESC LIMIT 1", (partner_tin, year)).fetchone()
        if r:
            return D(r["markup"])
    r = cx.execute(
        "SELECT markup FROM markup_rule WHERE year=? AND mxik IS NULL AND partner_tin IS NULL "
        "LIMIT 1", (year,)).fetchone()
    if r:
        return D(r["markup"])
    saved = get_setting(cx, "markup_by_year", {})
    if str(year) in saved:
        return D(saved[str(year)])
    return C.DEFAULT_MARKUP_BY_YEAR.get(year, C.DEFAULT_MARKUP)


def set_markup_for_year(cx, year, markup):
    saved = get_setting(cx, "markup_by_year", {})
    saved[str(year)] = str(markup)
    set_setting(cx, "markup_by_year", saved)


# ---------------------------------------------------------------------------
# Fayl dublikati
# ---------------------------------------------------------------------------
def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(131072), b""):
            h.update(chunk)
    return h.hexdigest()


def find_source_file(cx, sha):
    return cx.execute("SELECT * FROM source_file WHERE sha256=?", (sha,)).fetchone()


def add_source_file(cx, path, kind, sha=None):
    sha = sha or file_sha256(path)
    existing = find_source_file(cx, sha)
    if existing:
        return existing["id"], False
    cur = cx.execute(
        "INSERT INTO source_file(sha256,filename,full_path,kind,size_bytes,imported_at) "
        "VALUES(?,?,?,?,?,?)",
        (sha, os.path.basename(path), path, kind,
         os.path.getsize(path), datetime.datetime.now().isoformat(timespec="seconds")),
    )
    return cur.lastrowid, True


def update_source_counts(cx, sid, docs, lines):
    cx.execute("UPDATE source_file SET doc_count=?, line_count=? WHERE id=?",
               (docs, lines, sid))


# ---------------------------------------------------------------------------
# Mahsulot va alias
# ---------------------------------------------------------------------------
def find_product_by_norm(cx, norm):
    return cx.execute("SELECT * FROM product WHERE norm_name=? LIMIT 1", (norm,)).fetchone()


def find_product_by_alias(cx, norm):
    return cx.execute(
        "SELECT p.* FROM alias a JOIN product p ON p.id=a.product_id "
        "WHERE a.alias_norm=? LIMIT 1", (norm,)).fetchone()


def find_product_by_barcode(cx, barcode):
    if not barcode:
        return None
    return cx.execute(
        "SELECT * FROM product WHERE barcode=? AND barcode<>'' LIMIT 1",
        (str(barcode).strip(),)).fetchone()


def create_product(cx, canon_name, mxik=None, mxik_name=None, unit=None,
                   barcode=None, is_marked=None):
    norm = C.norm_name(canon_name)
    ex = find_product_by_norm(cx, norm)
    if ex:
        return ex["id"]
    cur = cx.execute(
        "INSERT INTO product(canon_name,norm_name,mxik,mxik_name,unit,barcode,"
        "is_marked,created_at) VALUES(?,?,?,?,?,?,?,?)",
        (canon_name, norm, mxik, mxik_name, unit, barcode,
         None if is_marked is None else int(is_marked),
         datetime.datetime.now().isoformat(timespec="seconds")),
    )
    pid = cur.lastrowid
    add_alias(cx, pid, canon_name, source="auto")
    return pid


def add_alias(cx, product_id, raw_name, source="manual"):
    norm = C.norm_name(raw_name)
    if not norm:
        return
    cx.execute(
        "INSERT INTO alias(product_id,alias_norm,alias_raw,source,created_at) "
        "VALUES(?,?,?,?,?) ON CONFLICT(alias_norm) DO NOTHING",
        (product_id, norm, raw_name, source,
         datetime.datetime.now().isoformat(timespec="seconds")),
    )


# ---------------------------------------------------------------------------
# Tekshiruv yozuvlari
# ---------------------------------------------------------------------------
def add_issue(cx, year, severity, code, message, ref_table=None, ref_id=None,
              detail=None):
    cx.execute(
        "INSERT INTO issue(year,severity,code,message,ref_table,ref_id,detail,created_at) "
        "VALUES(?,?,?,?,?,?,?,?)",
        (year, severity, code, message, ref_table, ref_id, detail,
         datetime.datetime.now().isoformat(timespec="seconds")),
    )


def clear_issues(cx, year=None, codes=None):
    sql = "DELETE FROM issue WHERE 1=1"
    args = []
    if year is not None:
        sql += " AND year=?"
        args.append(year)
    if codes:
        sql += " AND code IN (%s)" % ",".join("?" * len(codes))
        args.extend(codes)
    cx.execute(sql, args)


def issue_counts(cx, year=None):
    sql = ("SELECT severity, COUNT(*) n FROM issue WHERE resolved=0"
           + (" AND year=?" if year is not None else "") + " GROUP BY severity")
    rows = cx.execute(sql, (year,) if year is not None else ()).fetchall()
    return {r["severity"]: r["n"] for r in rows}


# ---------------------------------------------------------------------------
# Umumiy statistika
# ---------------------------------------------------------------------------
def available_years(cx, owner_tin=None):
    if owner_tin:
        rows = cx.execute(
            "SELECT DISTINCT doc_year y FROM document WHERE doc_year IS NOT NULL "
            "AND owner_tin=? ORDER BY y", (owner_tin,)).fetchall()
    else:
        rows = cx.execute(
            "SELECT DISTINCT doc_year y FROM document WHERE doc_year IS NOT NULL "
            "ORDER BY y").fetchall()
    return [r["y"] for r in rows]


def save_org(cx, tin, name=""):
    """Tashkilot (STIR) - nomi bo'sh bo'lsa keyin kelgan nom yoziladi."""
    if not tin:
        return
    cx.execute(
        "INSERT INTO org(tin, name) VALUES(?, ?) ON CONFLICT(tin) DO UPDATE SET "
        "name = CASE WHEN org.name = '' THEN excluded.name ELSE org.name END",
        (tin, (name or "").strip()))


def orgs(cx):
    """Bazadagi tashkilotlar: faktura/chek soni bilan, ko'pi oldin."""
    rows = cx.execute("""
        SELECT d.owner_tin tin, COALESCE(o.name, '') name,
               SUM(d.kind='kirim') kirim, SUM(d.kind='chiqim') chiqim
        FROM document d LEFT JOIN org o ON o.tin = d.owner_tin
        WHERE d.owner_tin IS NOT NULL AND d.owner_tin <> ''
        GROUP BY d.owner_tin ORDER BY COUNT(*) DESC""").fetchall()
    return [dict(r) for r in rows]


def org_name(cx, tin):
    r = cx.execute("SELECT name FROM org WHERE tin=?", (tin,)).fetchone()
    return (r["name"] if r else "") or ""


def unknown_owner_count(cx):
    """Qaysi tashkilotniki ekani noma'lum hujjatlar (eski faktura, fayli topilmagan)."""
    r = cx.execute("SELECT COUNT(*) FROM document WHERE owner_tin IS NULL OR owner_tin=''").fetchone()
    return r[0] if r else 0


def stats(cx):
    def one(sql, args=()):
        r = cx.execute(sql, args).fetchone()
        return r[0] if r else 0

    return {
        "files": one("SELECT COUNT(*) FROM source_file"),
        "docs_in": one("SELECT COUNT(*) FROM document WHERE kind='kirim'"),
        "docs_out": one("SELECT COUNT(*) FROM document WHERE kind='chiqim'"),
        "lines_in": one("SELECT COUNT(*) FROM doc_line WHERE kind='kirim'"),
        "lines_out": one("SELECT COUNT(*) FROM doc_line WHERE kind='chiqim'"),
        "products": one("SELECT COUNT(*) FROM product"),
        "unmatched": one(
            "SELECT COUNT(*) FROM doc_line WHERE product_id IS NULL"),
        "issues": one("SELECT COUNT(*) FROM issue WHERE resolved=0"),
        "years": available_years(cx),
    }


def reset_all(cx):
    """Hamma ma'lumotni o'chirish (sozlamalar va o'rganilgan aliaslar qoladi)."""
    for t in ("allocation", "stock_lot", "doc_line", "document",
              "source_file", "issue"):
        cx.execute("DELETE FROM %s" % t)
    cx.execute("VACUUM")
