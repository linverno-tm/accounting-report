# -*- coding: utf-8 -*-
"""
bh_parsers - uch xil manba formatini o'qish.

  1. FAKTURA/*.xls   -> aslida HTML (soliq.uz eksporti). Ustunlar soni
                        fayldan faylga o'zgaradi (10 / 11 / 12). Faqat
                        SARLAVHA MATNI bo'yicha o'qiladi.
  2. CHIQIM/*.xlsx   -> haqiqiy XLSX (checks-info). Ustunlar barqaror.
  3. eski .xls       -> haqiqiy BIFF8, mavjud hisobot (tarixni import qilish).

Har parser bir xil ko'rinish qaytaradi:
    {"documents": [ {..., "lines":[{...}]} ], "warnings": [str]}
"""

import os
import re
import datetime
from collections import OrderedDict, deque

import bh_config as C


# ===========================================================================
# 1. FAKTURA - HTML
# ===========================================================================
_RX_TITLE = re.compile(
    r"(?P<cdate>\d{1,2}\.\d{1,2}\.\d{4})\s*даги\s*(?P<cno>.+?)\s*-\s*сонли\s*шартномага"
    r"\s*(?P<ddate>\d{1,2}\.\d{1,2}\.\d{4})\s*даги\s*(?P<dno>.+?)\s*-\s*сонли",
    re.IGNORECASE | re.UNICODE,
)
_RX_TITLE_SIMPLE = re.compile(
    r"(?P<ddate>\d{1,2}\.\d{1,2}\.\d{4})\s*даги\s*(?P<dno>.+?)\s*-\s*сонли\s*Ҳисоб",
    re.IGNORECASE | re.UNICODE,
)


def _read_text(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "cp1251", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def _cell_text(td):
    t = td.get_text(" ", strip=True)
    return re.sub(r"\s+", " ", t.replace("\xa0", " ")).strip()


def _expand_row(tr):
    """colspan'ni hisobga olib satrni tekis ro'yxatga yoyadi."""
    out = []
    for td in tr.find_all(["td", "th"]):
        try:
            cs = int(td.get("colspan") or 1)
        except (TypeError, ValueError):
            cs = 1
        txt = _cell_text(td)
        out.append(txt)
        for _ in range(cs - 1):
            out.append(None)          # spanned bo'sh joy
    return out


def _flatten_header(rows, hi):
    """
    Sarlavhani tekis ro'yxatga aylantiradi.

    "ҚҚС" ustuni colspan=2 va ostida "Ставка" / "Сумма" turadi. Shuning
    uchun spanned joylarga pastki satr qiymatlari ketma-ket to'ldiriladi.
    """
    top = _expand_row(rows[hi])
    sub = []
    if hi + 1 < len(rows):
        sub = [x for x in _expand_row(rows[hi + 1]) if x]
    q = deque(sub)

    flat = []
    i = 0
    while i < len(top):
        cur = top[i]
        span = 1
        while i + span < len(top) and top[i + span] is None:
            span += 1
        if span == 1:
            flat.append(cur or "")
        else:
            for _ in range(span):
                s = q.popleft() if q else ""
                flat.append(("%s %s" % (cur or "", s)).strip())
        i += span
    return flat


def _map_header(flat):
    """Tekis sarlavhani maydon nomlariga bog'laydi: {field: index}."""
    idx = {}
    used = set()
    for j, h in enumerate(flat):
        hl = C.norm_name(h)
        if not hl:
            continue
        for field, keys in C.FAKTURA_HEADER_MAP:
            if field in idx:
                continue
            for k in keys:
                if C.norm_name(k) in hl:
                    if j in used:
                        continue
                    idx[field] = j
                    used.add(j)
                    break
            if field in idx:
                break
    return idx


def _split_mxik(text):
    """'09403001001000000 - Кровать (всех видов)' -> (kod, nom)."""
    if not text:
        return "", ""
    t = re.sub(r"\s+", " ", str(text)).strip()
    m = re.match(r"^\s*(\d[\d\s]{6,})\s*[-–—]\s*(.*)$", t)
    if m:
        return re.sub(r"\s", "", m.group(1)), m.group(2).strip()
    m = re.match(r"^\s*(\d[\d\s]{6,})\s*$", t)
    if m:
        return re.sub(r"\s", "", m.group(1)), ""
    return "", t


def _find_label(rows, *labels):
    """
    'Етказиб берувчи::' kabi yorliqdan keyingi qiymatni topadi.

    Faktura HTML'ida bitta katta "blob" satr ham bor - unda hamma yorliq va
    qiymat aralash yotadi. Shuning uchun:
      - yorliq katakchasi QISQA bo'lishi kerak (blob emas)
      - qiymat o'zi yorliq bo'lib qolmasligi kerak (':' bilan tugamasin)
    """
    want = [C.norm_name(x) for x in labels]
    for tr in rows:
        cells = [c for c in _expand_row(tr) if c is not None]
        for i, c in enumerate(cells[:-1]):
            if not c or len(c) > 80:          # blob satrni tashlab ketish
                continue
            cn = C.norm_name(c)
            if not any(cn.startswith(w) for w in want):
                continue
            v = (cells[i + 1] or "").strip()
            if not v or v.endswith(":") or len(v) > 300:
                continue
            return v
    return ""


def parse_faktura_html(path):
    from bs4 import BeautifulSoup

    warnings = []
    soup = BeautifulSoup(_read_text(path), "lxml")
    table = soup.find("table")
    if table is None:
        return {"documents": [], "warnings": ["%s: jadval topilmadi" % os.path.basename(path)]}
    rows = table.find_all("tr")

    # --- sarlavha satrini topish ---
    hi = None
    for i, tr in enumerate(rows):
        cells = [c for c in _expand_row(tr) if c is not None]
        if cells and cells[0] == "№" and len(cells) > 5:
            hi = i
            break
    if hi is None:
        return {"documents": [], "warnings": ["%s: ustun sarlavhasi topilmadi" % os.path.basename(path)]}

    flat = _flatten_header(rows, hi)
    idx = _map_header(flat)
    for need in ("name", "qty", "price"):
        if need not in idx:
            return {"documents": [],
                    "warnings": ["%s: '%s' ustuni topilmadi (ustunlar: %s)"
                                 % (os.path.basename(path), need, " | ".join(flat))]}

    # --- hujjat sarlavhasi ---
    title = _cell_text(rows[0]) if rows else ""
    doc_no = doc_date = contract_no = contract_date = ""
    m = _RX_TITLE.search(title) or _RX_TITLE_SIMPLE.search(title)
    if m:
        g = m.groupdict()
        doc_no = re.sub(r"\s+", " ", g.get("dno", "")).strip()
        doc_date = g.get("ddate", "")
        contract_no = re.sub(r"\s+", " ", g.get("cno", "") or "").strip()
        contract_date = g.get("cdate", "") or ""
    if not doc_date:
        # fayl nomidan: ..._21_07_2026_дан
        fm = re.search(r"_(\d{2})_(\d{2})_(\d{4})_", os.path.basename(path))
        if fm:
            doc_date = "%s.%s.%s" % fm.groups()
            warnings.append("%s: sana sarlavhadan o'qilmadi, fayl nomidan olindi"
                            % os.path.basename(path))

    d_date = C.parse_date(doc_date)
    supplier = _find_label(rows, "етказиб берувчи")
    supplier_tin = _find_label(rows, "етказиб берувчининг стир")
    buyer = _find_label(rows, "сотиб олувчи")
    buyer_tin = _find_label(rows, "сотиб олувчининг стир")

    # --- satrlar ---
    lines = []
    ncols = len(flat)
    for tr in rows[hi + 1:]:
        cells = _expand_row(tr)
        cells = [("" if c is None else c) for c in cells]
        if not cells:
            continue
        first = cells[0].strip()
        if not re.match(r"^\d+$", first):
            continue
        if len(cells) < ncols - 2:
            continue
        # "1 2 3 4..." raqamlash satrini tashlab ketish
        seq = [c.strip() for c in cells[1:6]]
        if seq[:4] == ["1", "2", "3", "4"]:
            continue

        def g(field, default=""):
            j = idx.get(field)
            if j is None or j >= len(cells):
                return default
            return cells[j].strip()

        qty = C.parse_number(g("qty"))
        price = C.parse_number(g("price"))
        if qty is None and price is None:
            continue

        mxik_code, mxik_name = _split_mxik(g("mxik"))
        unit_raw = g("unit")
        is_marked, mark_code = C.parse_marking(g("marking"))
        amount_net = C.parse_number(g("amount_net"))
        vat_amount = C.parse_number(g("vat_amount"))
        amount_gross = C.parse_number(g("amount_gross"))

        # Faktura'da Нарҳ = BIRLIK narxi, QQS'SIZ.
        # (50 x 298 135,71 = 14 906 785,71 = "Етказиб бериш қиймати" - tekshirildi)
        if amount_net is None and qty is not None and price is not None:
            amount_net = qty * price
        if vat_amount is None and amount_net is not None:
            vat_amount = C.money(amount_net) * C.VAT_RATE
        if amount_gross is None and amount_net is not None:
            amount_gross = C.money(amount_net) + C.money(vat_amount or 0)

        lines.append({
            "line_no": int(first),
            "raw_name": g("name"),
            "norm_name": C.norm_name(g("name")),
            "mxik": mxik_code,
            "mxik_name": mxik_name,
            "barcode": re.sub(r"\D", "", g("barcode")) or "",
            "marking_code": mark_code,
            "is_marked": is_marked,
            "unit_raw": unit_raw,
            "unit": C.norm_unit(unit_raw),
            "packaging": C.unit_packaging(unit_raw),
            "qty": C.qty(qty or 0),
            "unit_price_net": C.money(price or 0),
            "amount_net": C.money(amount_net or 0),
            "vat_amount": C.money(vat_amount or 0),
            "amount_gross": C.money(amount_gross or 0),
            "unit_price_gross": None,
            "line_date": d_date.isoformat() if d_date else None,
        })

    if not lines:
        warnings.append("%s: birorta tovar satri topilmadi" % os.path.basename(path))

    doc = {
        "kind": "kirim",
        "doc_key": "faktura|%s|%s|%s" % (supplier_tin or "?", doc_no or "?",
                                         d_date.isoformat() if d_date else "?"),
        "doc_no": doc_no,
        "doc_date": d_date.isoformat() if d_date else None,
        "doc_year": d_date.year if d_date else None,
        "contract_no": contract_no,
        "contract_date": (C.parse_date(contract_date).isoformat()
                          if C.parse_date(contract_date) else None),
        "partner_name": supplier,
        "partner_tin": supplier_tin,
        "pos_id": None,
        "check_type": None,
        "is_return": 0,
        "total_net": sum((l["amount_net"] for l in lines), C.money(0)),
        "total_vat": sum((l["vat_amount"] for l in lines), C.money(0)),
        "total_gross": sum((l["amount_gross"] for l in lines), C.money(0)),
        "lines": lines,
        "buyer_name": buyer,
        "buyer_tin": buyer_tin,
    }
    return {"documents": [doc], "warnings": warnings}


# ===========================================================================
# 2. CHIQIM - checks-info XLSX
# ===========================================================================
def _sheets_from_xlsx(path):
    """
    Har varaqning satrlari: [(varaq_nomi, [satr, ...]), ...]

    !!! read_only rejimida openpyxl varaqning <dimension> yozuviga ishonadi
    va undan keyingi satrlarni JIMGINA tashlab ketadi. Eksport qiluvchi
    dasturlar bu yozuvni ko'pincha noto'g'ri qo'yadi (masalan A1:T8) -
    natijada 12 oylik chek faylidan faqat birinchi 1-2 oy o'qilardi.
    reset_dimensions() varaqni oxirigacha o'qishga majburlaydi.

    Barcha varaqlar o'qiladi: oylar alohida varaqlarda bo'lishi mumkin.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        out = []
        for ws in wb.worksheets:
            try:
                ws.reset_dimensions()
            except AttributeError:
                pass
            out.append((ws.title, list(ws.iter_rows(values_only=True))))
        return out
    finally:
        wb.close()


def _is_checks_header(row):
    vals = [C.norm_name(v) for v in (row or ())]
    return any("маҳсулот номи" in v or "махсулот номи" in v for v in vals) and \
        any("чек санаси" in v for v in vals)


def _find_checks_header(rows):
    """Sarlavha satrining indeksi."""
    for i, row in enumerate(rows[:30]):
        if _is_checks_header(row):
            return i
    return None


def parse_checks_xlsx(path):
    warnings = []
    fname = os.path.basename(path)
    sheets = _sheets_from_xlsx(path)
    if not any(rows for _t, rows in sheets):
        return {"documents": [], "warnings": ["%s: bo'sh fayl" % fname]}

    # Sarlavhasi bor har varaqdan sarlavhadan keyingi satrlar yig'iladi
    data_rows = []
    found = False
    for _title, rows in sheets:
        hi = _find_checks_header(rows)
        if hi is None:
            continue
        found = True
        data_rows.extend(rows[hi + 1:])
    if not found:
        return {"documents": [],
                "warnings": ["%s: 'Чеклар рўйхати' sarlavhasi topilmadi" % fname]}

    F = C.CHECKS_COLUMNS
    checks = OrderedDict()
    no_date = 0

    for row in data_rows:
        if row is None or row[0] is None or str(row[0]).strip() == "":
            continue
        # Bir nechta eksport bitta varaqqa qo'shib yuborilgan bo'lsa,
        # sarlavha o'rtada takrorlanadi
        if _is_checks_header(row):
            continue
        r = {}
        for j, field in enumerate(F):
            r[field] = row[j] if j < len(row) else None

        dt = C.parse_date(r["check_dt"])
        if dt is None:
            no_date += 1
        pos_id = str(r["pos_id"] or "").strip()
        check_no = str(r["check_no"] or "").strip()
        if check_no.endswith(".0"):
            check_no = check_no[:-2]
        dt_raw = str(r["check_dt"] or "").strip()

        key = "chek|%s|%s|%s" % (pos_id, check_no, dt_raw)
        if key not in checks:
            tin = str(r["tin"] or "").strip()
            if tin.endswith(".0"):
                tin = tin[:-2]
            is_ret = C.is_return_check(r["check_type"])
            checks[key] = {
                "kind": "chiqim",
                "doc_key": key,
                "doc_no": check_no,
                "doc_date": dt.isoformat() if dt else None,
                "doc_year": dt.year if dt else None,
                "contract_no": None,
                "contract_date": None,
                "partner_name": None,
                "partner_tin": tin,
                "pos_id": pos_id,
                "check_type": str(r["check_type"] or "").strip(),
                "is_return": 1 if is_ret else 0,
                # !!! Chek jami - HAR SATRDA TAKRORLANADI, shuning uchun
                # faqat birinchi satrdan olinadi.
                "check_cash": C.money(C.parse_number(r["check_cash"]) or 0),
                "check_card": C.money(C.parse_number(r["check_card"]) or 0),
                "check_vat": C.money(C.parse_number(r["check_vat"]) or 0),
                "lines": [],
                "_dt_raw": dt_raw,
            }

        doc = checks[key]
        sign = -1 if doc["is_return"] else 1

        qty = C.parse_number(r["qty"]) or 0
        # !!! "Нархи" ustuni NARX EMAS, SATR SUMMASI (QQS ichida).
        gross = C.parse_number(r["line_gross"]) or 0
        vat = C.parse_number(r["line_vat"])
        if vat is None:
            vat = C.vat_from_gross(gross)
        else:
            # Manbadagi qiymat BIRLIK QQS emas, satr QQS si (tekshirildi:
            # 84000 x 12/112 = 9000 - aynan mos keldi)
            vat = C.money(vat)

        pid_raw = str(r["product_id"] or "").strip()
        m = re.search(r"-(\d+)$", pid_raw)
        line_no = int(m.group(1)) + 1 if m else len(doc["lines"]) + 1

        raw_name = str(r["name"] or "").strip()
        is_marked, mark_code = C.parse_marking(r["marking"])
        unit_raw = str(r["unit_code"] or "").strip()

        q = C.qty(qty) * sign
        g = C.money(gross) * sign
        v = C.money(vat) * sign
        net = C.money(g) - C.money(v)

        doc["lines"].append({
            "line_no": line_no,
            "raw_name": raw_name,
            "norm_name": C.norm_name(raw_name),
            "mxik": str(r["mxik"] or "").strip(),
            "mxik_name": "",
            "barcode": re.sub(r"\D", "", str(r["barcode"] or "")) or "",
            "marking_code": mark_code,
            "is_marked": is_marked,
            "unit_raw": unit_raw,
            "unit": "",          # kassada birlik kodi, nomi emas
            "packaging": "",
            "qty": q,
            "unit_price_net": (C.money(net / q) if q else C.money(0)),
            "amount_net": C.money(net),
            "vat_amount": v,
            "amount_gross": g,
            "unit_price_gross": (C.money(g / q) if q else C.money(0)),
            "line_date": dt.isoformat() if dt else None,
            "product_ext_id": pid_raw,
        })

    docs = []
    for doc in checks.values():
        lines_sum = sum((l["amount_gross"] for l in doc["lines"]), C.money(0))
        check_total = C.money(doc["check_cash"]) + C.money(doc["check_card"])
        doc["total_gross"] = lines_sum
        doc["total_vat"] = sum((l["vat_amount"] for l in doc["lines"]), C.money(0))
        doc["total_net"] = C.money(lines_sum) - C.money(doc["total_vat"])
        doc["_check_total"] = check_total
        doc["_lines_sum"] = lines_sum
        doc.pop("_dt_raw", None)
        docs.append(doc)

    if no_date:
        # Sanasiz chek hech qaysi yilga tushmaydi - hisobotda ko'rinmaydi
        warnings.append("%s: %d satrda chek sanasi o'qilmadi - bu satrlar hisobotga "
                        "tushmaydi" % (fname, no_date))
    return {"documents": docs, "warnings": warnings}


# ===========================================================================
# 3. Eski hisobot (.xls, BIFF8) - tarixni import qilish
# ===========================================================================
def _xl_cell(sh, wb, r, c):
    if c >= sh.ncols or r >= sh.nrows:
        return None
    cl = sh.cell(r, c)
    if cl.ctype == 0 or cl.ctype == 6:
        return None
    if cl.ctype == 5:
        return None                      # formula xatosi (#DIV/0! va h.k.)
    if cl.ctype == 3:
        import xlrd
        try:
            return xlrd.xldate.xldate_as_datetime(cl.value, wb.datemode).date()
        except Exception:
            return None
    return cl.value


def parse_legacy_report(path):
    """
    Mavjud "КАМЕРАЛ ТЕКШИРУВЛАР" faylidan:
      - "касса YYYY" varaqlari -> chiqim hujjatlari
      - "YYYY ТХ" varaqlari    -> kirim satrlari (tannarx bilan)
    """
    import xlrd

    warnings = []
    wb = xlrd.open_workbook(path, formatting_info=False)
    docs = []

    for sh in wb.sheets():
        nm = sh.name.strip()
        ym = re.search(r"(20\d{2})", nm)
        year = int(ym.group(1)) if ym else None
        low = C.norm_name(nm)

        # --- kassa varaqlari ---
        if "касса" in low:
            checks = OrderedDict()
            for r in range(1, sh.nrows):
                pid = _xl_cell(sh, wb, r, 0)
                if pid is None or not str(pid).strip():
                    continue
                vals = [_xl_cell(sh, wb, r, c) for c in range(min(20, sh.ncols))]
                rec = dict(zip(C.CHECKS_COLUMNS, vals))
                dt = C.parse_date(rec.get("check_dt"))
                pos_id = str(rec.get("pos_id") or "").strip()
                cno = str(rec.get("check_no") or "").strip()
                if cno.endswith(".0"):
                    cno = cno[:-2]
                key = "chek|%s|%s|%s" % (pos_id, cno, str(rec.get("check_dt") or ""))
                if key not in checks:
                    checks[key] = {
                        "kind": "chiqim", "doc_key": key, "doc_no": cno,
                        "doc_date": dt.isoformat() if dt else None,
                        "doc_year": dt.year if dt else year,
                        "contract_no": None, "contract_date": None,
                        "partner_name": None,
                        "partner_tin": str(rec.get("tin") or "").strip(),
                        "pos_id": pos_id,
                        "check_type": str(rec.get("check_type") or "").strip(),
                        "is_return": 1 if C.is_return_check(rec.get("check_type")) else 0,
                        "check_cash": C.money(C.parse_number(rec.get("check_cash")) or 0),
                        "check_card": C.money(C.parse_number(rec.get("check_card")) or 0),
                        "check_vat": C.money(C.parse_number(rec.get("check_vat")) or 0),
                        "lines": [],
                    }
                doc = checks[key]
                sign = -1 if doc["is_return"] else 1
                q = C.qty(C.parse_number(rec.get("qty")) or 0) * sign
                g = C.money(C.parse_number(rec.get("line_gross")) or 0) * sign
                v = C.money(C.parse_number(rec.get("line_vat")) or 0) * sign
                if not v:
                    v = C.vat_from_gross(g)
                name = str(rec.get("name") or "").strip()
                im, mc = C.parse_marking(rec.get("marking"))
                doc["lines"].append({
                    "line_no": len(doc["lines"]) + 1,
                    "raw_name": name, "norm_name": C.norm_name(name),
                    "mxik": str(rec.get("mxik") or "").strip(), "mxik_name": "",
                    "barcode": re.sub(r"\D", "", str(rec.get("barcode") or "")) or "",
                    "marking_code": mc, "is_marked": im,
                    "unit_raw": str(rec.get("unit_code") or "").strip(),
                    "unit": "", "packaging": "",
                    "qty": q,
                    "unit_price_net": C.money((g - v) / q) if q else C.money(0),
                    "amount_net": C.money(g) - C.money(v),
                    "vat_amount": v, "amount_gross": g,
                    "unit_price_gross": C.money(g / q) if q else C.money(0),
                    "line_date": dt.isoformat() if dt else None,
                })
            for d in checks.values():
                d["total_gross"] = sum((l["amount_gross"] for l in d["lines"]), C.money(0))
                d["total_vat"] = sum((l["vat_amount"] for l in d["lines"]), C.money(0))
                d["total_net"] = C.money(d["total_gross"]) - C.money(d["total_vat"])
                d["_check_total"] = C.money(d["check_cash"]) + C.money(d["check_card"])
                d["_lines_sum"] = d["total_gross"]
                docs.append(d)

        # --- ТХ varaqlari (kirim) ---
        elif "тх" in low.split() or low.endswith("тх"):
            lines = []
            for r in range(4, sh.nrows):
                name = _xl_cell(sh, wb, r, 2)
                if not (isinstance(name, str) and name.strip()):
                    continue
                d = _xl_cell(sh, wb, r, 1)
                ld = d if isinstance(d, datetime.date) else C.parse_date(d)
                q_in = C.parse_number(_xl_cell(sh, wb, r, 10))
                p_in = C.parse_number(_xl_cell(sh, wb, r, 11))
                if not q_in or q_in == 0:
                    continue
                mxik_code, mxik_name = _split_mxik(_xl_cell(sh, wb, r, 4))
                unit_raw = str(_xl_cell(sh, wb, r, 5) or "")
                im, mc = C.parse_marking(_xl_cell(sh, wb, r, 3))
                cost = C.money(p_in or 0)
                if cost < 0:
                    warnings.append("%s r%d: manfiy tannarx (%s) - musbatga aylantirildi"
                                    % (nm, r + 1, cost))
                    cost = -cost
                amount = C.money(cost * C.qty(q_in))
                lines.append({
                    "line_no": len(lines) + 1,
                    "raw_name": name.strip(), "norm_name": C.norm_name(name),
                    "mxik": mxik_code, "mxik_name": mxik_name,
                    "barcode": "", "marking_code": mc, "is_marked": im,
                    "unit_raw": unit_raw, "unit": C.norm_unit(unit_raw),
                    "packaging": C.unit_packaging(unit_raw),
                    "qty": C.qty(q_in),
                    "unit_price_net": cost,
                    "amount_net": amount,
                    "vat_amount": C.money(amount * C.VAT_RATE),
                    "amount_gross": C.money(amount * (1 + C.VAT_RATE)),
                    "unit_price_gross": None,
                    "line_date": ld.isoformat() if ld else None,
                    "_legacy_sale_price": C.parse_number(_xl_cell(sh, wb, r, 6)),
                })
            if lines:
                docs.append({
                    "kind": "kirim",
                    "doc_key": "legacy|%s" % nm,
                    "doc_no": nm, "doc_date": None, "doc_year": year,
                    "contract_no": None, "contract_date": None,
                    "partner_name": "Eski hisobotdan", "partner_tin": None,
                    "pos_id": None, "check_type": None, "is_return": 0,
                    "total_net": sum((l["amount_net"] for l in lines), C.money(0)),
                    "total_vat": sum((l["vat_amount"] for l in lines), C.money(0)),
                    "total_gross": sum((l["amount_gross"] for l in lines), C.money(0)),
                    "lines": lines,
                })
    return {"documents": docs, "warnings": warnings}


# ===========================================================================
# Fayl turini aniqlash
# ===========================================================================
def sniff(path):
    """Faylning HAQIQIY turini bayt boshidan aniqlaydi (kengaytmaga ishonmaydi)."""
    try:
        head = open(path, "rb").read(512)
    except Exception:
        return "nomalum"
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"                       # OLE2 / BIFF8
    if head[:2] == b"PK":
        return "xlsx"
    low = head.lstrip()[:120].lower()
    if low.startswith(b"<html") or low.startswith(b"<!doctype") or b"<table" in low:
        return "html"
    if b"<?xml" in low and b"workbook" in head.lower():
        return "xmlss"
    return "nomalum"


def guess_kind(path):
    """kirim | chiqim | legacy | nomalum"""
    base = os.path.basename(path).lower()
    kind = sniff(path)
    if "checks-info" in base or "чеклар" in base:
        return "chiqim"
    if "фактура" in base or "faktura" in base or "hisob" in base or "ҳисоб" in base:
        return "kirim"
    if kind == "html":
        return "kirim"
    if kind == "xlsx":
        return "chiqim"
    if kind == "xls":
        return "legacy"
    return "nomalum"


def parse_any(path, kind=None):
    """Fayl turini aniqlab mos parserni chaqiradi."""
    kind = kind or guess_kind(path)
    fmt = sniff(path)
    name = os.path.basename(path)

    if kind == "kirim":
        if fmt == "html":
            return parse_faktura_html(path)
        if fmt == "xlsx":
            return {"documents": [], "warnings": [
                "%s: kirim deb belgilandi, lekin XLSX. Hozircha faqat "
                "soliq.uz HTML fakturasi qo'llab-quvvatlanadi." % name]}
        return {"documents": [], "warnings": ["%s: format tanilmadi (%s)" % (name, fmt)]}

    if kind == "chiqim":
        if fmt == "xlsx":
            return parse_checks_xlsx(path)
        if fmt == "xls":
            return parse_legacy_report(path)
        return {"documents": [], "warnings": ["%s: format tanilmadi (%s)" % (name, fmt)]}

    if kind == "legacy":
        if fmt == "xls":
            return parse_legacy_report(path)
        return {"documents": [], "warnings": ["%s: eski hisobot .xls bo'lishi kerak" % name]}

    return {"documents": [], "warnings": ["%s: turini aniqlab bo'lmadi" % name]}


SUPPORTED_EXT = (".xls", ".xlsx", ".xlsm", ".htm", ".html")


def scan_folder(folder, recursive=True):
    """Papkadan mos fayllarni topadi."""
    out = []
    if not os.path.isdir(folder):
        return out
    if recursive:
        for root, _dirs, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(SUPPORTED_EXT) and not f.startswith("~$"):
                    out.append(os.path.join(root, f))
    else:
        for f in sorted(os.listdir(folder)):
            p = os.path.join(folder, f)
            if os.path.isfile(p) and f.lower().endswith(SUPPORTED_EXT) \
                    and not f.startswith("~$"):
                out.append(p)
    return sorted(out)
