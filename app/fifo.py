# -*- coding: utf-8 -*-
"""
bh_fifo - import, FIFO taqsimot, tekshiruv.

Eski hisobotdagi ikkita asosiy nuqson shu modulda hal qilinadi:

  1. YILLARARO QOLDIQ O'TMAGAN. Tekshirilgan: 2023-2026 varaqlarining
     hammasida "давр бошига қолдиқ" nol. Bu yerda partiya (stock_lot)
     yillardan o'tib ketaveradi, hisobot esa har yil uchun kesimni
     hisoblab beradi.

  2. MANFIY QOLDIQ. Eski faylda 2023 r25: kirim 2 dona, chiqim 5 dona,
     qoldiq -3. 2026 jami qoldiq -6 372.92 dona. Bu yerda FIFO omborda
     bori qadar yechadi, yetmagani ALOHIDA belgilanadi va buxgalterga
     ko'rsatiladi - jimgina manfiyga tushmaydi.
"""

import os
import datetime
from decimal import Decimal

import bh_config as C
import bh_db as DB
import bh_matching as M

ZERO = Decimal("0")

# SQL: tashkilot (STIR) bo'yicha filtr; owner_tin=None - hammasi. Ikki marta uzatiladi.
OWN = "(? IS NULL OR %s.owner_tin = ?)"


# ===========================================================================
# Import
# ===========================================================================
def import_parsed(cx, engine, parsed, source_file_id, on_issue=None, quiet_dups=False):
    """
    Parserdan kelgan hujjatlarni bazaga yozadi.
    (yangi_hujjat, otkazib_yuborilgan, satr) qaytaradi.

    quiet_dups=True - allaqachon bor hujjat uchun "takror" yozuvi qo'shilmaydi
    (eski faylni qayta o'qishda minglab keraksiz yozuv chiqmasligi uchun).
    """
    added = skipped = nlines = 0

    # Ulanish avtokommit rejimida (isolation_level=None), ya'ni har INSERT
    # alohida tranzaksiya bo'lib ketadi. Minglab satrda bu eng katta
    # sekinlik - shuning uchun hammasi bitta tranzaksiyaga o'raladi.
    cx.execute("BEGIN")
    try:
        added, skipped, nlines = _import_docs(cx, engine, parsed, source_file_id,
                                              quiet_dups)
        cx.execute("COMMIT")
    except Exception:
        cx.execute("ROLLBACK")
        raise
    return added, skipped, nlines


# ---------------------------------------------------------------------------
# Eski chek fayllarini qayta o'qish
#
# 1.3.0 gacha chek fayli (checks-info) to'liq o'qilmasdi: faqat birinchi
# varaq, u ham faylning <dimension> yozuvigacha - amalda ko'pincha 1-2 oy.
# Fayl sha256 bo'yicha "kiritilgan" deb belgilangani uchun uni qayta
# tashlash ham yordam bermasdi. Shu sababli har eski chek fayli BIR MARTA
# qayta o'qiladi. Mavjud cheklar doc_key bo'yicha o'tkazib yuboriladi,
# faqat yetishmagan oylar qo'shiladi - ikki marta hisoblanmaydi.
# ---------------------------------------------------------------------------
def _reparse_key():
    return "checks_reparsed_v%d" % C.CHECKS_PARSER_VER


def _reparsed(cx):
    return set(DB.get_setting(cx, _reparse_key(), []) or [])


def needs_reparse(cx, sf):
    """Bu source_file yozuvi eski parser bilan o'qilganmi?"""
    return sf is not None and sf["kind"] == "chiqim" and \
        sf["sha256"] not in _reparsed(cx)


def _mark_reparsed(cx, shas):
    done = _reparsed(cx)
    done.update(shas)
    DB.set_setting(cx, _reparse_key(), sorted(done))


def reparse_file(cx, engine, sf, path):
    """Bitta eski chek faylini qayta o'qiydi. (yangi_hujjat, satr, ogohlantirishlar)."""
    import bh_parsers as P

    r = P.parse_any(path, sf["kind"])
    a, _s, n = import_parsed(cx, engine, r, sf["id"], quiet_dups=True)
    cx.execute("UPDATE source_file SET doc_count=doc_count+?, line_count=line_count+?, "
               "full_path=? WHERE id=?", (a, n, path, sf["id"]))
    _mark_reparsed(cx, [sf["sha256"]])
    return a, n, r["warnings"]


def backfill_owners(cx, progress=None):
    """
    Eski bazadagi fakturalarga egasini (sotib oluvchi STIRi) yozadi - fayl asl
    joyida tursa qayta o'qiladi. 1.5.0 gacha bu STIR saqlanmagan, shuning uchun
    ikki do'kon fakturalari bitta hisobotga aralashib ketardi.
    {"fixed": hujjat, "missing": topilmagan fayl} qaytaradi.
    """
    import bh_parsers as P

    rows = cx.execute("""
        SELECT DISTINCT sf.id, sf.full_path, sf.sha256, sf.filename
        FROM source_file sf JOIN document d ON d.source_file_id = sf.id
        WHERE d.kind = 'kirim' AND (d.owner_tin IS NULL OR d.owner_tin = '')""").fetchall()
    fixed = missing = 0
    for i, sf in enumerate(rows):
        path = sf["full_path"] or ""
        try:
            ok = os.path.isfile(path) and DB.file_sha256(path) == sf["sha256"]
        except OSError:
            ok = False
        if not ok:
            missing += 1
            continue
        if progress:
            progress(i, len(rows), "Tashkilotlar ajratilmoqda: %d / %d" % (i + 1, len(rows)))
        for doc in P.parse_any(path, "kirim")["documents"]:
            own = doc.get("owner_tin")
            if not own:
                continue
            DB.save_org(cx, own, doc.get("buyer_name") or "")
            cur = cx.execute(
                "UPDATE document SET owner_tin=? WHERE doc_key=? "
                "AND (owner_tin IS NULL OR owner_tin='')", (own, doc["doc_key"]))
            fixed += cur.rowcount
    return {"fixed": fixed, "missing": missing}


def reparse_old_checks(cx, engine, progress=None):
    """
    Bazadagi barcha eski chek fayllarini (asl joyida turgan bo'lsa) qayta
    o'qiydi. {"files", "docs", "lines", "missing": [fayl nomi]} qaytaradi.
    """
    done = _reparsed(cx)
    todo = [sf for sf in cx.execute(
        "SELECT * FROM source_file WHERE kind='chiqim' ORDER BY id").fetchall()
        if sf["sha256"] not in done]
    res = {"files": 0, "docs": 0, "lines": 0, "missing": []}
    for i, sf in enumerate(todo):
        path = sf["full_path"] or ""
        ok = False
        try:
            ok = os.path.isfile(path) and DB.file_sha256(path) == sf["sha256"]
        except OSError:
            ok = False
        if not ok:
            # Fayl ko'chirilgan / o'chirilgan - buxgalter qayta tashlasa
            # do_import uni shu yerga qaytaradi
            res["missing"].append(sf["filename"])
            continue
        if progress:
            progress(i, len(todo), "eski chek fayli: %s" % sf["filename"])
        a, n, _w = reparse_file(cx, engine, sf, path)
        res["files"] += 1
        res["docs"] += a
        res["lines"] += n
    return res


def _import_docs(cx, engine, parsed, source_file_id, quiet_dups=False):
    added = skipped = nlines = 0

    for doc in parsed["documents"]:
        owner = doc.get("owner_tin") or None
        if owner:
            DB.save_org(cx, owner, doc.get("buyer_name") if doc["kind"] == "kirim" else "")
        ex = cx.execute("SELECT id, owner_tin FROM document WHERE doc_key=?",
                        (doc["doc_key"],)).fetchone()
        if ex:
            skipped += 1
            if owner and not ex["owner_tin"]:
                # eski bazadagi hujjatga egasi yoziladi (fayl qayta qo'shilganda)
                cx.execute("UPDATE document SET owner_tin=? WHERE id=?", (owner, ex["id"]))
            if not quiet_dups:
                DB.add_issue(cx, doc.get("doc_year"), C.SEVERITY_INFO, "duplicate_doc",
                             "Hujjat allaqachon kiritilgan: %s" % (doc.get("doc_no") or "?"),
                             "document", ex["id"])
            continue

        cur = cx.execute(
            "INSERT INTO document(source_file_id,kind,doc_key,doc_no,doc_date,doc_year,"
            "contract_no,contract_date,partner_name,partner_tin,pos_id,check_type,"
            "total_net,total_vat,total_gross,is_return,owner_tin) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (source_file_id, doc["kind"], doc["doc_key"], doc.get("doc_no"),
             doc.get("doc_date"), doc.get("doc_year"), doc.get("contract_no"),
             doc.get("contract_date"), doc.get("partner_name"), doc.get("partner_tin"),
             doc.get("pos_id"), doc.get("check_type"),
             str(doc.get("total_net") or 0), str(doc.get("total_vat") or 0),
             str(doc.get("total_gross") or 0), int(doc.get("is_return") or 0), owner),
        )
        did = cur.lastrowid
        added += 1

        # Chek jami satrlar yig'indisiga to'g'ri keldimi?
        if doc["kind"] == "chiqim" and "_check_total" in doc:
            ct = C.money(doc["_check_total"])
            ls = C.money(doc["_lines_sum"])
            if ct and abs(ct - ls) > Decimal("1"):
                DB.add_issue(
                    cx, doc.get("doc_year"), C.SEVERITY_WARN, "check_total_mismatch",
                    "Chek №%s: satrlar yig'indisi %s, chek jami %s (farq %s)"
                    % (doc.get("doc_no"), ls, ct, ls - ct), "document", did)

        for ln in doc["lines"]:
            if doc["kind"] == "kirim":
                pid = engine.ensure_product(ln)
                meth, score = "kirim", 1.0
            else:
                pid, meth, score, _ = engine.match(ln)
                if pid and score < M.AUTO_ACCEPT_SCORE:
                    pid = None
                if not pid:
                    meth, score = "", 0.0

            cx.execute(
                "INSERT INTO doc_line(document_id,kind,line_no,raw_name,norm_name,mxik,"
                "mxik_name,barcode,marking_code,is_marked,unit_raw,unit,packaging,qty,"
                "unit_price_net,amount_net,vat_amount,amount_gross,unit_price_gross,"
                "product_id,match_method,match_score,line_date) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (did, doc["kind"], ln.get("line_no"), ln.get("raw_name"),
                 ln.get("norm_name"), ln.get("mxik"), ln.get("mxik_name"),
                 ln.get("barcode"),
                 ln.get("marking_code"),
                 None if ln.get("is_marked") is None else int(ln["is_marked"]),
                 ln.get("unit_raw"), ln.get("unit"), ln.get("packaging"),
                 str(ln.get("qty") or 0), str(ln.get("unit_price_net") or 0),
                 str(ln.get("amount_net") or 0), str(ln.get("vat_amount") or 0),
                 str(ln.get("amount_gross") or 0),
                 None if ln.get("unit_price_gross") is None
                 else str(ln["unit_price_gross"]),
                 pid, meth, float(score),
                 ln.get("line_date") or doc.get("doc_date")),
            )
            nlines += 1

    return added, skipped, nlines


# ===========================================================================
# Ombor qayta qurish (FIFO)
# ===========================================================================
def rebuild_stock(cx, progress=None):
    """
    stock_lot va allocation jadvallarini noldan qayta quradi.

    Tartib: BARCHA kirimlar va chiqimlar sana bo'yicha aralash tartibda
    qayta o'ynatiladi - shuning uchun yil chegarasi hech qanday ahamiyatga
    ega emas, qoldiq o'z-o'zidan o'tadi.
    """
    cx.execute("BEGIN")
    try:
        res = _rebuild(cx, progress)
        cx.execute("COMMIT")
    except Exception:
        cx.execute("ROLLBACK")
        raise
    return res


def _rebuild(cx, progress=None):
    """
    Ombor ikki bosqichda qayta quriladi.

    NEGA IKKI BOSQICH. Ilgari kirim va chiqim BITTA zanjirda, sana bo'yicha
    aralash o'ynatilardi. Natijada sotuv sanasi faktura sanasidan oldin
    bo'lsa (buxgalter fakturani keyinroq kiritsa yoki soliq.uz fakturani
    keyin bersa), o'sha paytda ochiq partiya topilmasdi va sotuv "omborda
    yo'q" deb belgilanardi - holbuki tovar omborda bor edi. Partiyadan
    hech narsa yechilmagani uchun qoldiq ham kamaymasdi.

    Endi:
      1-bosqich - barcha kirim partiyalari yaratiladi;
      2-bosqich - sotuvlar sana tartibida yechiladi:
            a) sotuv sanasida ochiq bo'lgan partiyalardan (haqiqiy FIFO);
            b) topilmasa - SHU YIL ichidagi keyingi kirimdan (ogohlantirish
               bilan: sana tartibi teskari, lekin tovar davr ichida bor);
            c) butun davrda ham bo'lmasa - "omborda yo'q" (kirim hujjati
               umuman yetishmaydi).

    (b) qoidasi buxgalterning talabi: "товар чиқим қилинганда омборда
    мавжуд товар қолдиғидан чиқим қилиш керак".
    """
    cx.execute("DELETE FROM allocation")
    cx.execute("DELETE FROM stock_lot")
    DB.clear_issues(cx, codes=["no_stock", "negative_cost", "negative_qty",
                               "unmatched_sale", "unit_mismatch", "late_receipt"])

    rows = cx.execute("""
        SELECT l.id, l.kind, l.product_id, l.qty, l.unit_price_net,
               l.amount_net, l.amount_gross, l.raw_name, l.unit,
               COALESCE(l.line_date, d.doc_date) AS dt,
               d.doc_year, d.doc_no, d.partner_tin, d.is_return,
               COALESCE(d.owner_tin, '') AS owner
        FROM doc_line l JOIN document d ON d.id = l.document_id
        WHERE l.product_id IS NOT NULL
        ORDER BY COALESCE(l.line_date, d.doc_date) IS NULL,
                 COALESCE(l.line_date, d.doc_date),
                 CASE l.kind WHEN 'kirim' THEN 0 ELSE 1 END,
                 l.id
    """).fetchall()

    total = len(rows)
    # Bir tovarning bir necha kartochkasi (bh_matching.product_groups) bitta
    # zaxira sifatida yuritiladi: sotuv qaysi kartochkaga bog'langan bo'lsa
    # ham, guruhdagi eng eski partiyadan yechiladi.
    grp = M.load_product_groups(cx)
    # guruh ildizi -> partiyalar [ [lot_id, qty_left, unit_cost, lot_date], ... ]
    open_lots = {}
    last_cost = {}
    n_alloc = n_short = n_late = 0
    allocs = []          # bitta executemany uchun to'planadi
    shorts = []          # ombor yetmagan satrlar
    touched = {}         # lot_id -> oxirgi qty_left

    # =======================================================================
    # 1-BOSQICH: kirim partiyalari
    # =======================================================================
    sales = []
    for r in rows:
        if r["kind"] != "kirim":
            sales.append(r)
            continue

        pid = r["product_id"]
        q = DB.D(r["qty"])
        dt = r["dt"] or ""
        year = r["doc_year"]

        if q <= 0:
            DB.add_issue(cx, year, C.SEVERITY_WARN, "negative_qty",
                         "Kirim miqdori musbat emas: %s (%s)" % (q, r["raw_name"]),
                         "doc_line", r["id"])
            continue
        cost = DB.D(r["unit_price_net"])
        if cost < 0:
            DB.add_issue(cx, year, C.SEVERITY_ERROR, "negative_cost",
                         "Manfiy tannarx %s: %s" % (cost, r["raw_name"]),
                         "doc_line", r["id"])
            cost = -cost
        if cost == 0:
            DB.add_issue(cx, year, C.SEVERITY_WARN, "missing_price",
                         "Kirim narxi nol: %s" % r["raw_name"], "doc_line", r["id"])
        markup = DB.get_markup(cx, year or 0, None, r["partner_tin"])
        cur = cx.execute(
            "INSERT INTO stock_lot(product_id,doc_line_id,lot_date,lot_year,"
            "qty_in,qty_left,unit_cost,markup,is_opening,source_year) "
            "VALUES(?,?,?,?,?,?,?,?,0,?)",
            (pid, r["id"], dt, year, str(q), str(q), str(cost), str(markup), year))
        # Zaxira kaliti: (tashkilot, tovar guruhi) - bir do'kon sotuvi boshqasining
        # kirimidan yechilmaydi
        g = (r["owner"], grp.get(pid, pid))
        open_lots.setdefault(g, []).append([cur.lastrowid, q, cost, dt])
        last_cost[g] = cost

    # =======================================================================
    # 2-BOSQICH: sotuvlar
    # =======================================================================
    def _consume(lots, need, sale_line_id, pid, dt, year, only_future_of_year=None):
        """
        Partiyalardan yechadi va qolgan ehtiyojni qaytaradi.

        only_future_of_year=None -> sotuv sanasida OCHIQ partiyalar
        only_future_of_year=YYYY -> sotuvdan KEYIN kelgan, lekin shu yildagi
        """
        nonlocal n_alloc
        for lot in lots:
            if need <= 0:
                break
            if lot[1] <= 0:
                continue
            ld = lot[3] or ""
            if only_future_of_year is None:
                # sanasi yo'q partiya har doim mavjud deb qaraladi
                if ld and dt and ld > dt:
                    continue
            else:
                if not (ld and dt and ld > dt and ld[:4] == only_future_of_year):
                    continue
            take = lot[1] if lot[1] <= need else need
            allocs.append((sale_line_id, lot[0], pid, dt, year, str(C.qty(take)),
                           str(lot[2]), str(C.money(take * lot[2]))))
            n_alloc += 1
            lot[1] -= take
            need -= take
            touched[lot[0]] = lot[1]        # qty_left oxirida bir marta yoziladi
        return need

    for i, r in enumerate(sales):
        pid = r["product_id"]
        g = (r["owner"], grp.get(pid, pid))
        q = DB.D(r["qty"])
        dt = r["dt"] or ""
        year = r["doc_year"]

        if q == 0:
            continue
        if q < 0:
            # Qaytarish: tovar omborga qaytadi (eng oxirgi tannarx bilan)
            back = -q
            cost = last_cost.get(g) or ZERO
            cur = cx.execute(
                "INSERT INTO stock_lot(product_id,doc_line_id,lot_date,lot_year,"
                "qty_in,qty_left,unit_cost,markup,is_opening,source_year) "
                "VALUES(?,?,?,?,?,?,?,?,0,?)",
                (pid, r["id"], dt, year, str(back), str(back), str(cost),
                 str(DB.get_markup(cx, year or 0)), year))
            open_lots.setdefault(g, []).insert(0, [cur.lastrowid, back, cost, dt])
            continue

        lots = open_lots.get(g, [])

        # a) sotuv sanasida ochiq partiyalardan
        need = _consume(lots, q, r["id"], pid, dt, year)

        # b) topilmasa - shu yil ichidagi keyingi kirimdan
        if need > 0:
            ysale = (dt[:4] if dt else None) or (str(year) if year else None)
            if ysale:
                before = need
                need = _consume(lots, need, r["id"], pid, dt, year,
                                only_future_of_year=ysale)
                if need < before:
                    n_late += 1
                    DB.add_issue(
                        cx, year, C.SEVERITY_WARN, "late_receipt",
                        "%s: %s dona %s da sotilgan, kirim hujjati kechroq sana "
                        "bilan kiritilgan - davr ichidagi kirimdan yechildi"
                        % ((r["raw_name"] or "?")[:60], C.qty(before - need),
                           dt or "?"), "doc_line", r["id"])

        # c) butun davrda ham yo'q
        if need > 0:
            est = last_cost.get(g)
            if est is None:
                gross = DB.D(r["amount_gross"])
                mk = DB.get_markup(cx, year or 0)
                est = C.money(C.net_from_gross(gross / q if q else 0) / (1 + mk)) \
                    if q else ZERO
            shorts.append((r["id"], pid, dt, year, str(C.qty(need)), str(est),
                           str(C.money(need * est))))
            n_short += 1
            DB.add_issue(
                cx, year, C.SEVERITY_ERROR, "no_stock",
                "%s: %s dona sotilgan, lekin omborda yo'q (kirim hujjati yetishmaydi). "
                "Tannarx taxminiy: %s" % ((r["raw_name"] or "?")[:60], C.qty(need), est),
                "doc_line", r["id"])

        if progress and (i % 500 == 0 or i == len(sales) - 1):
            progress(i + 1, len(sales))

    # --- to'plangan yozuvlar bir zarbda ---
    if allocs:
        cx.executemany(
            "INSERT INTO allocation(sale_line_id,stock_lot_id,product_id,"
            "alloc_date,alloc_year,qty,unit_cost,cost_amount,shortfall) "
            "VALUES(?,?,?,?,?,?,?,?,0)", allocs)
    if shorts:
        cx.executemany(
            "INSERT INTO allocation(sale_line_id,stock_lot_id,product_id,"
            "alloc_date,alloc_year,qty,unit_cost,cost_amount,shortfall) "
            "VALUES(?,NULL,?,?,?,?,?,?,1)", shorts)
    if touched:
        cx.executemany("UPDATE stock_lot SET qty_left=? WHERE id=?",
                       [(str(C.qty(q)), lid) for lid, q in touched.items()])

    return {"lines": total, "allocations": n_alloc, "shortfalls": n_short,
            "late": n_late}


# ===========================================================================
# Yil kesimi - hisobot uchun
# ===========================================================================
def year_rows(cx, year, owner_tin=None):
    """
    Hisobotning "YYYY ТХ" varag'i uchun satrlar.

    Har satr = bitta kirim partiyasi (eski fayldagi kabi: bir mahsulot
    bir necha marta uchraydi, har kirim sanasi bilan).

    Har partiya uchun yil kesimi:
        boshiga qoldiq = qty_in (agar partiya yildan OLDIN kelgan bo'lsa)
                         minus yil boshigacha yechilgani
        kirim          = qty_in (agar partiya SHU yilda kelgan bo'lsa)
        chiqim         = shu yil ichida yechilgani
        oxiriga qoldiq = boshiga + kirim - chiqim

    owner_tin berilsa - faqat shu tashkilot (STIR) hujjatlari.
    """
    y0 = "%d-01-01" % year
    own = (owner_tin, owner_tin)
    y1 = "%d-12-31" % year

    lots = cx.execute("""
        SELECT s.id, s.product_id, s.doc_line_id, s.lot_date, s.lot_year,
               s.qty_in, s.unit_cost, s.markup,
               p.canon_name, p.mxik, p.mxik_name, p.unit AS punit, p.markup AS pmarkup,
               l.raw_name, l.unit_raw, l.unit, l.packaging, l.is_marked,
               l.marking_code, l.mxik AS lmxik, l.mxik_name AS lmxik_name
        FROM stock_lot s
        JOIN product p ON p.id = s.product_id
        LEFT JOIN doc_line l ON l.id = s.doc_line_id
        LEFT JOIN document dd ON dd.id = l.document_id
        WHERE s.lot_date <= ? AND %s
        ORDER BY s.lot_date, s.id
    """ % (OWN % "dd"), (y1,) + own).fetchall()

    # partiya -> (yil boshigacha yechilgan, shu yil yechilgan)
    before = {}
    during = {}
    for r in cx.execute(
            "SELECT stock_lot_id, alloc_date, SUM(CAST(qty AS REAL)) q "
            "FROM allocation WHERE stock_lot_id IS NOT NULL GROUP BY stock_lot_id, alloc_date"):
        lid, d, q = r["stock_lot_id"], r["alloc_date"] or "", DB.D(r["q"])
        if d < y0:
            before[lid] = before.get(lid, ZERO) + q
        elif d <= y1:
            during[lid] = during.get(lid, ZERO) + q

    # shu yilda partiyadan sotilgan summa (sotish narxida, QQS bilan)
    sale_gross = {}
    for r in cx.execute("""
            SELECT a.stock_lot_id lot, a.qty aq, l.unit_price_gross upg,
                   l.qty lq, l.amount_gross ag
            FROM allocation a JOIN doc_line l ON l.id = a.sale_line_id
            WHERE a.stock_lot_id IS NOT NULL AND a.alloc_date >= ? AND a.alloc_date <= ?
        """, (y0, y1)):
        aq = DB.D(r["aq"])
        upg = DB.D(r["upg"]) if r["upg"] is not None else None
        if upg is None:
            lq = DB.D(r["lq"])
            upg = (DB.D(r["ag"]) / lq) if lq else ZERO
        sale_gross[r["lot"]] = sale_gross.get(r["lot"], ZERO) + C.money(aq * upg)

    default_markup = DB.get_markup(cx, year)
    out = []
    for r in lots:
        lot_id = r["id"]
        qin_total = DB.D(r["qty_in"])
        cost = DB.D(r["unit_cost"])
        b = before.get(lot_id, ZERO)
        d = during.get(lot_id, ZERO)

        is_this_year = (r["lot_date"] or "")[:4] == str(year)
        if is_this_year:
            open_q = ZERO
            in_q = qin_total
        else:
            open_q = qin_total - b
            in_q = ZERO
            if open_q <= 0 and d == 0:
                continue          # yil boshiga ham, ichida ham hech narsa yo'q

        out_q = d
        close_q = open_q + in_q - out_q
        if open_q == 0 and in_q == 0 and out_q == 0:
            continue

        markup = DB.D(r["pmarkup"]) if r["pmarkup"] else (
            DB.D(r["markup"]) if r["markup"] else default_markup)
        sale_price = C.sale_price_from_cost(cost, markup)

        unit_raw = r["unit_raw"] or r["punit"] or ""
        mxik = r["lmxik"] or r["mxik"] or ""
        mxik_name = r["lmxik_name"] or r["mxik_name"] or ""

        out.append({
            "lot_id": lot_id,
            "product_id": r["product_id"],
            "date": r["lot_date"],
            "name": r["raw_name"] or r["canon_name"],
            "is_marked": r["is_marked"],
            "marking_code": r["marking_code"] or "",
            "mxik": mxik,
            "mxik_name": mxik_name,
            "unit": unit_raw,
            "sale_price": sale_price,
            "markup": markup,
            "cost": cost,
            "open_qty": C.qty(open_q), "open_sum": C.money(open_q * cost),
            "in_qty": C.qty(in_q), "in_sum": C.money(in_q * cost),
            "out_qty": C.qty(out_q), "out_sum": C.money(out_q * cost),
            "close_qty": C.qty(close_q), "close_sum": C.money(close_q * cost),
            "sale_qty": C.qty(out_q),
            "sale_sum": C.money(sale_gross.get(lot_id, ZERO)),
        })

    # Omborsiz sotilganlar (kirim hujjati yo'q) - alohida satr sifatida
    for r in cx.execute("""
            SELECT a.product_id, a.alloc_date, SUM(CAST(a.qty AS REAL)) q,
                   AVG(CAST(a.unit_cost AS REAL)) c, p.canon_name, p.mxik,
                   p.mxik_name, p.unit,
                   SUM(CAST(a.qty AS REAL) * CAST(l.amount_gross AS REAL)
                       / NULLIF(CAST(l.qty AS REAL), 0)) g
            FROM allocation a JOIN product p ON p.id=a.product_id
            JOIN doc_line l ON l.id = a.sale_line_id
            JOIN document dd ON dd.id = l.document_id
            WHERE a.shortfall=1 AND a.alloc_date >= ? AND a.alloc_date <= ? AND %s
            GROUP BY a.product_id
        """ % (OWN % "dd"), (y0, y1) + own):
        q = DB.D(r["q"])
        cost = DB.D(r["c"])
        # sotuv summasi - chekdagi haqiqiy summa (taxminiy narxdan emas)
        gross = C.money(DB.D(r["g"])) if r["g"] is not None else \
            C.money(q * C.sale_price_from_cost(cost, default_markup))
        out.append({
            "lot_id": None, "product_id": r["product_id"], "date": r["alloc_date"],
            "name": r["canon_name"], "is_marked": None, "marking_code": "",
            "mxik": r["mxik"] or "", "mxik_name": r["mxik_name"] or "",
            "unit": r["unit"] or "",
            "sale_price": C.sale_price_from_cost(cost, default_markup),
            "markup": default_markup, "cost": cost,
            "open_qty": ZERO, "open_sum": ZERO,
            "in_qty": ZERO, "in_sum": ZERO,
            "out_qty": C.qty(q), "out_sum": C.money(q * cost),
            "close_qty": C.qty(-q), "close_sum": C.money(-q * cost),
            "sale_qty": C.qty(q),
            "sale_sum": gross,
            "shortfall": True,
        })

    # Hech qaysi kirimga bog'lanmagan sotuvlar. Ilgari ular ТХ ga umuman
    # tushmasdi - kirim fakturasi yo'q tovar (masalan "MEBEL") sotilgan oy
    # hisobotdan butunlay yo'qolardi. Endi nom + oy bo'yicha qizil satr:
    # miqdor va sotuv summasi aniq, tannarx sotuv narxidan taxminiy.
    # Buxgalter "Bog'lash" oynasida bog'lasa, satr oddiy partiyaga o'tadi.
    for r in cx.execute("""
            SELECT l.norm_name, MIN(l.raw_name) raw_name, MIN(l.mxik) mxik,
                   MIN(l.unit_raw) unit_raw, MIN(COALESCE(l.line_date, d.doc_date)) dt,
                   SUM(CAST(l.qty AS REAL)) q, SUM(CAST(l.amount_gross AS REAL)) g,
                   MAX(l.is_marked) is_marked
            FROM doc_line l JOIN document d ON d.id = l.document_id
            WHERE l.kind='chiqim' AND l.product_id IS NULL
              AND COALESCE(l.line_date, d.doc_date) >= ?
              AND COALESCE(l.line_date, d.doc_date) <= ? AND %s
            GROUP BY l.norm_name, substr(COALESCE(l.line_date, d.doc_date), 1, 7)
        """ % (OWN % "d"), (y0, y1 + "~") + own):
        q = C.qty(DB.D(r["q"]))
        if q == 0:
            continue
        gross = C.money(DB.D(r["g"]))
        sale_price = C.money(gross / q)
        cost = C.money(C.net_from_gross(sale_price) / (1 + default_markup))
        out.append({
            "lot_id": None, "product_id": None, "date": r["dt"],
            "name": r["raw_name"], "is_marked": r["is_marked"], "marking_code": "",
            "mxik": r["mxik"] or "", "mxik_name": "", "unit": r["unit_raw"] or "",
            "sale_price": sale_price, "markup": default_markup, "cost": cost,
            "open_qty": ZERO, "open_sum": ZERO,
            "in_qty": ZERO, "in_sum": ZERO,
            "out_qty": q, "out_sum": C.money(q * cost),
            "close_qty": -q, "close_sum": C.money(-q * cost),
            "sale_qty": q, "sale_sum": gross,
            "shortfall": True, "unmatched": True,
        })

    out.sort(key=lambda x: (x["date"] or "", x["name"] or ""))
    return out


def year_totals(rows):
    def s(k):
        return sum((DB.D(r[k]) for r in rows), ZERO)
    return {
        "open_sum": C.money(s("open_sum")), "in_sum": C.money(s("in_sum")),
        "out_sum": C.money(s("out_sum")), "close_sum": C.money(s("close_sum")),
        "sale_sum": C.money(s("sale_sum")),
        "open_qty": C.qty(s("open_qty")), "in_qty": C.qty(s("in_qty")),
        "out_qty": C.qty(s("out_qty")), "close_qty": C.qty(s("close_qty")),
        "profit": C.money(C.net_from_gross(s("sale_sum")) - s("out_sum")),
        "rows": len(rows),
    }


# ===========================================================================
# Tekshiruv
# ===========================================================================
def validate_year(cx, year, owner_tin=None):
    """Hisobot oldidan mantiqiy tekshiruvlar."""
    DB.clear_issues(cx, year=year, codes=["wrong_year", "bad_date", "unit_mismatch"])

    for r in cx.execute(
            "SELECT id, doc_no, doc_date, doc_year, kind FROM document "
            "WHERE doc_date IS NULL OR doc_date=''"):
        DB.add_issue(cx, year, C.SEVERITY_WARN, "bad_date",
                     "Hujjat sanasi o'qilmadi: %s (%s)" % (r["doc_no"], r["kind"]),
                     "document", r["id"])

    for r in cx.execute(
            "SELECT d.id, d.doc_no, d.doc_date, d.doc_year FROM document d "
            "WHERE d.doc_year IS NOT NULL AND substr(d.doc_date,1,4) <> CAST(d.doc_year AS TEXT)"):
        DB.add_issue(cx, year, C.SEVERITY_WARN, "wrong_year",
                     "Hujjat №%s sanasi %s, lekin %d yilga yozilgan"
                     % (r["doc_no"], r["doc_date"], r["doc_year"]), "document", r["id"])

    for r in cx.execute("""
            SELECT p.id, p.canon_name, COUNT(DISTINCT l.unit) n,
                   GROUP_CONCAT(DISTINCT l.unit) units
            FROM doc_line l JOIN product p ON p.id=l.product_id
            WHERE l.unit IS NOT NULL AND l.unit<>''
            GROUP BY p.id HAVING n > 1"""):
        DB.add_issue(cx, year, C.SEVERITY_WARN, "unit_mismatch",
                     "%s: o'lchov birligi har xil (%s)"
                     % (r["canon_name"][:60], r["units"]), "product", r["id"])

    n = cx.execute(
        "SELECT COUNT(*) c FROM doc_line l JOIN document d ON d.id=l.document_id "
        "WHERE l.kind='chiqim' AND l.product_id IS NULL AND d.doc_year=? AND " + OWN % "d",
        (year, owner_tin, owner_tin)).fetchone()["c"]
    if n:
        # detail = STIR: Xatolar varag'i shu tashkilotnikini ko'rsatadi
        DB.add_issue(cx, year, C.SEVERITY_ERROR, "unmatched_sale",
                     "%d ta sotuv satri kirimga bog'lanmagan - tannarx aniqlanmaydi" % n,
                     detail=owner_tin)
    return DB.issue_counts(cx, year)
