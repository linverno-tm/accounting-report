# -*- coding: utf-8 -*-
"""
bh_matching - sotuv satrini kirim mahsulotiga bog'lash.

Bu ilovaning eng qiyin qismi. Sabab (haqiqiy ma'lumotda o'lchandi):
  - kassa mahsulot nomini 63 BELGIDA kesadi
  - aniq nom mos kelishi atigi 40/89 (45%)
  - MXIK kalit bo'la olmaydi: 08517001001000000 -> 21 xil telefon

Yechim - ishonch darajasi bo'yicha zina. Har bosqich o'zining ballini
beradi, past balli natija buxgalterga tasdiqlash uchun chiqadi.
Buxgalter bir marta tasdiqlagan juftlik alias sifatida saqlanadi va
boshqa hech qachon so'ralmaydi.
"""

import re
import difflib
from collections import defaultdict

import bh_config as C
import bh_db as DB


# Bosqich -> (usul nomi, ball, avtomatik qabul qilinadimi)
METHOD_BARCODE = ("barcode", 1.00, True)
METHOD_MARKING = ("marking", 1.00, True)
METHOD_ALIAS = ("alias", 0.99, True)
METHOD_EXACT = ("name", 0.98, True)
METHOD_PREFIX = ("prefix", 0.94, True)
METHOD_FUZZY_MXIK = ("fuzzy+mxik", 0.0, True)     # ball hisoblanadi
METHOD_FUZZY = ("fuzzy", 0.0, False)
# MXIK ostida bitta tovar (yoki bitta tovarning kartochkalari) bo'lsa -
# avtomatik. Aks holda sotuv umuman bog'lanmay, hisobotdan tushib qolardi.
# Umumiy MXIK (08517001001000000 -> 21 telefon) bu qoidaga tushmaydi,
# chunki u yerda tovar bitta emas.
METHOD_MXIK_UNIQUE = ("mxik", 0.90, True)
METHOD_MANUAL = ("manual", 1.00, True)

AUTO_ACCEPT_SCORE = 0.88      # shundan yuqori - avtomatik
SUGGEST_SCORE = 0.55          # shundan yuqori - taklif sifatida ko'rsatiladi


def _tokens(norm):
    return [t for t in norm.split() if len(t) > 1]


# ---------------------------------------------------------------------------
# Bir xil tovarning bir nechta kartochkasi
#
# Faktura nomi ham kesilib keladi, shuning uchun bitta tovar ikki kartochka
# bo'lib qoladi (haqiqiy ma'lumotdan):
#   "Велосипеды двухколёсные, ... с номинальной мощностью 240"
#   "Велосипеды двухколёсные, ... с номинальной мощностью 240 Вт"
#   "Мотороллер электрический, ... для перевозки грузов"
#   "Мотороллер электрический, ... для перевозки грузов, DRONGO"
# Kassa nomi ikkalasiga ham mos keladi, moslashtirish birini tanlay
# olmasdi va sotuv bog'lanmay qolardi - butun oy hisobotdan tushib qolardi.
#
# Qoida: MXIK bir xil VA qisqa nom uzun nomning boshlanishi bo'lsa - bitta
# tovar. Guruh ildizi = eng kichik id (birinchi kirim).
# ---------------------------------------------------------------------------
SAME_ITEM_MIN_LEN = 25


def same_item(n1, n2):
    a, b = (n1, n2) if len(n1) <= len(n2) else (n2, n1)
    return len(a) >= SAME_ITEM_MIN_LEN and b.startswith(a)


def product_groups(products):
    """{product_id: guruh_ildizi}. products = {pid: {norm_name, mxik, ...}}."""
    parent = {pid: pid for pid in products}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    by_mxik = defaultdict(list)
    for pid, p in products.items():
        m = (p.get("mxik") or "").strip()
        if m:
            by_mxik[m].append(pid)
    for pids in by_mxik.values():
        pids.sort()
        for i, a in enumerate(pids):
            na = products[a].get("norm_name") or ""
            for b in pids[i + 1:]:
                if same_item(na, products[b].get("norm_name") or ""):
                    ra, rb = find(a), find(b)
                    if ra != rb:
                        parent[max(ra, rb)] = min(ra, rb)
    return {pid: find(pid) for pid in products}


def load_product_groups(cx):
    return product_groups({r["id"]: dict(r) for r in
                           cx.execute("SELECT id, norm_name, mxik FROM product")})


def _model_codes(norm):
    """
    Nomdagi model belgilarini ajratadi: 'k304pro', 'sh-hd-9898', '1103-t'.
    Aynan shular mahsulotni farqlaydi (MXIK emas).
    """
    out = set()
    for t in re.findall(r"[a-z0-9][a-z0-9\-]{2,}", norm):
        if re.search(r"\d", t) and re.search(r"[a-z]", t):
            out.add(t.replace("-", ""))
        elif t.isdigit() and len(t) >= 3:
            out.add(t)
    return out


class MatchEngine:
    """Mahsulot indekslari + moslashtirish zinasi."""

    def __init__(self, cx):
        self.cx = cx
        self.reload()

    # -- indekslar ------------------------------------------------------
    def reload(self):
        cx = self.cx
        self.by_norm = {}
        self.by_barcode = {}
        self.by_mxik = defaultdict(list)
        self.products = {}
        self.norm_list = []

        for r in cx.execute("SELECT * FROM product"):
            pid = r["id"]
            self.products[pid] = dict(r)
            n = r["norm_name"] or ""
            if n:
                self.by_norm.setdefault(n, pid)
                self.norm_list.append((n, pid))
            b = (r["barcode"] or "").strip()
            if b:
                self.by_barcode.setdefault(b, pid)
            m = (r["mxik"] or "").strip()
            if m:
                self.by_mxik[m].append(pid)

        self.by_alias = {}
        for r in cx.execute("SELECT alias_norm, product_id FROM alias"):
            self.by_alias[r["alias_norm"]] = r["product_id"]

        # 63-belgili prefiks indeksi (kassa kesgan nomlar uchun)
        self.by_prefix = defaultdict(set)
        for n, pid in self.norm_list:
            self.by_prefix[n[:40]].add(pid)
        for a, pid in self.by_alias.items():
            self.by_prefix[a[:40]].add(pid)

        self._model_idx = defaultdict(set)
        for n, pid in self.norm_list:
            for mc in _model_codes(n):
                self._model_idx[mc].add(pid)

        self._groups = None

    def group_of(self, pid):
        """Mahsulot guruhining ildizi (bir xil tovar kartochkalari)."""
        if self._groups is None:
            self._groups = product_groups(self.products)
        return self._groups.get(pid, pid)

    def _single_item(self, pids):
        """pids hammasi bitta tovarmi? Bo'lsa - guruh ildizi, aks holda None."""
        roots = {self.group_of(p) for p in pids}
        return roots.pop() if len(roots) == 1 else None

    # -- mahsulot yaratish (faqat KIRIM satridan) -----------------------
    def ensure_product(self, line):
        """
        Kirim satri uchun mahsulot kartochkasi. Kirim - tannarxning yagona
        manbai, shuning uchun mahsulot kartochkasi shu yerdan tug'iladi.
        """
        cx = self.cx
        norm = line.get("norm_name") or C.norm_name(line.get("raw_name"))
        bar = (line.get("barcode") or "").strip()

        if bar and bar in self.by_barcode:
            pid = self.by_barcode[bar]
            self._learn(pid, line)
            return pid
        if norm in self.by_alias:
            pid = self.by_alias[norm]
            self._learn(pid, line)
            return pid
        if norm in self.by_norm:
            pid = self.by_norm[norm]
            self._learn(pid, line)
            return pid

        pid = DB.create_product(
            cx,
            canon_name=(line.get("raw_name") or "").strip(),
            mxik=line.get("mxik") or None,
            mxik_name=line.get("mxik_name") or None,
            unit=line.get("unit") or None,
            barcode=bar or None,
            is_marked=line.get("is_marked"),
        )
        row = cx.execute("SELECT * FROM product WHERE id=?", (pid,)).fetchone()
        self.products[pid] = dict(row)
        if norm:
            self.by_norm.setdefault(norm, pid)
            self.norm_list.append((norm, pid))
            self.by_prefix[norm[:40]].add(pid)
            for mc in _model_codes(norm):
                self._model_idx[mc].add(pid)
        if bar:
            self.by_barcode.setdefault(bar, pid)
        if line.get("mxik"):
            self.by_mxik[line["mxik"]].append(pid)
        self.by_alias[norm] = pid
        self._groups = None
        return pid

    def _learn(self, pid, line):
        """Mahsulotga yangi ma'lumot kelsa (barkod, MXIK) to'ldirib qo'yadi."""
        cx = self.cx
        p = self.products.get(pid) or {}
        upd = {}
        if not (p.get("barcode") or "").strip() and (line.get("barcode") or "").strip():
            upd["barcode"] = line["barcode"].strip()
        if not (p.get("mxik") or "").strip() and (line.get("mxik") or "").strip():
            upd["mxik"] = line["mxik"].strip()
        if not (p.get("unit") or "").strip() and (line.get("unit") or "").strip():
            upd["unit"] = line["unit"].strip()
        if upd:
            cx.execute(
                "UPDATE product SET %s WHERE id=?" % ",".join("%s=?" % k for k in upd),
                list(upd.values()) + [pid],
            )
            p.update(upd)
            if "barcode" in upd:
                self.by_barcode.setdefault(upd["barcode"], pid)
            if "mxik" in upd:
                self.by_mxik[upd["mxik"]].append(pid)
        norm = line.get("norm_name") or C.norm_name(line.get("raw_name"))
        if norm and norm not in self.by_alias:
            DB.add_alias(cx, pid, line.get("raw_name") or "", source="auto")
            self.by_alias[norm] = pid
            self.by_prefix[norm[:40]].add(pid)

    # -- moslashtirish zinasi (CHIQIM satri uchun) -----------------------
    def match(self, line):
        """
        (product_id | None, usul, ball, takliflar) qaytaradi.
        takliflar = [(product_id, ball, sabab)] - qo'lda tanlash uchun.
        """
        raw = (line.get("raw_name") or "").strip()
        norm = line.get("norm_name") or C.norm_name(raw)
        bar = (line.get("barcode") or "").strip()
        mark = (line.get("marking_code") or "").strip()
        mxik = (line.get("mxik") or "").strip()

        # 1) Barkod
        if bar and bar in self.by_barcode:
            return self.by_barcode[bar], METHOD_BARCODE[0], METHOD_BARCODE[1], []

        # 2) Marking kodi
        if mark:
            r = self.cx.execute(
                "SELECT product_id FROM doc_line WHERE marking_code=? AND kind='kirim' "
                "AND product_id IS NOT NULL LIMIT 1", (mark,)).fetchone()
            if r:
                return r["product_id"], METHOD_MARKING[0], METHOD_MARKING[1], []

        # 3) O'rganilgan alias
        if norm and norm in self.by_alias:
            return self.by_alias[norm], METHOD_ALIAS[0], METHOD_ALIAS[1], []

        # 4) Aniq nom
        if norm and norm in self.by_norm:
            return self.by_norm[norm], METHOD_EXACT[0], METHOD_EXACT[1], []

        # 5) Kesilgan nom (kassa 63 belgida kesadi)
        if raw and len(raw) >= C.POS_NAME_MAX_LEN - 2:
            hits = self._prefix_candidates(norm)
            if len(hits) == 1:
                return hits[0], METHOD_PREFIX[0], METHOD_PREFIX[1], []
            root = self._single_item(hits) if len(hits) > 1 else None
            if root is not None:
                # bir tovarning bir necha kartochkasi - ombor ularni birga yuritadi
                return root, METHOD_PREFIX[0], METHOD_PREFIX[1], []
            if len(hits) > 1:
                # bir nechta nomzod - modelga qarab tanlanadi
                best = self._rank(norm, mxik, hits)
                if best and best[0][1] >= AUTO_ACCEPT_SCORE:
                    return best[0][0], METHOD_PREFIX[0], best[0][1], best[1:6]
                return None, "", 0.0, best[:6]

        # 6/7) Fuzzy - avval MXIK ichida, keyin model kodi, keyin umumiy
        cands = set()
        if mxik and mxik in self.by_mxik:
            cands.update(self.by_mxik[mxik])
        for mc in _model_codes(norm):
            cands.update(self._model_idx.get(mc, ()))
        if not cands:
            cands.update(self._prefix_candidates(norm))

        ranked = self._rank(norm, mxik, cands)
        if ranked and ranked[0][1] >= AUTO_ACCEPT_SCORE:
            meth = METHOD_FUZZY_MXIK[0] if mxik and ranked[0][0] in self.by_mxik.get(mxik, []) \
                else METHOD_FUZZY[0]
            return ranked[0][0], meth, ranked[0][1], ranked[1:6]

        # 8) MXIK ostida bitta tovar bo'lsa - avtomatik
        if mxik and self.by_mxik.get(mxik):
            root = self._single_item(self.by_mxik[mxik])
            if root is not None:
                return root, METHOD_MXIK_UNIQUE[0], METHOD_MXIK_UNIQUE[1], ranked[:6]

        return None, "", 0.0, [r for r in ranked if r[1] >= SUGGEST_SCORE][:6]

    def _prefix_candidates(self, norm):
        if not norm:
            return []
        key = norm[:40]
        hits = set(self.by_prefix.get(key, ()))
        if not hits:
            # qisqaroq prefiks bilan qayta urinish
            for n, pid in self.norm_list:
                if n.startswith(norm[: max(12, len(norm) - 3)]):
                    hits.add(pid)
        else:
            hits = {p for p in hits
                    if C.norm_name(self.products[p]["canon_name"]).startswith(
                        norm[: max(10, len(norm) - 2)])}
        return sorted(hits)

    def _rank(self, norm, mxik, cands):
        """Nomzodlarni ball bo'yicha tartiblaydi."""
        out = []
        ntok = set(_tokens(norm))
        nmodels = _model_codes(norm)
        for pid in cands:
            p = self.products.get(pid)
            if not p:
                continue
            pn = p["norm_name"] or ""
            floor = 0.0
            score = difflib.SequenceMatcher(None, norm, pn).ratio()
            reason = "o'xshashlik %.0f%%" % (score * 100)
            # to'liq nom kassa nomi bilan boshlansa - kuchli dalil
            if pn.startswith(norm[: max(10, len(norm) - 2)]) and len(norm) > 20:
                score = max(score, 0.95)
                reason = "kesilgan nomga mos"
            # aksincha: kassa nomi = faktura nomi + rang/izoh
            # ("ЛДСП 2750*1830*16мм Серый графит" <- "ЛДСП 2750*1830*16мм")
            elif mxik and (p["mxik"] or "") == mxik and len(pn) >= 12 and \
                    norm.startswith(pn + " "):
                floor = 0.93           # so'z o'xshashligi bilan aralashtirilmaydi
                reason = "faktura nomi + izoh"
            # model kodi aynan uchrasa - juda kuchli dalil
            pmodels = _model_codes(pn)
            if nmodels and pmodels and (nmodels & pmodels):
                score = min(1.0, score + 0.12)
                reason += ", model mos (%s)" % ",".join(sorted(nmodels & pmodels)[:2])
            elif nmodels and pmodels and not (nmodels & pmodels):
                score -= 0.25          # model boshqa - deyarli aniq boshqa tovar
                reason += ", MODEL BOSHQA"
                floor = 0.0
            if mxik and (p["mxik"] or "") == mxik:
                score = min(1.0, score + 0.05)
                reason += ", MXIK mos"
            elif mxik and (p["mxik"] or ""):
                score -= 0.05
            ptok = set(_tokens(pn))
            if ntok and ptok:
                jac = len(ntok & ptok) / float(len(ntok | ptok))
                score = (score * 0.75) + (jac * 0.25)
            score = max(score, floor)
            out.append((pid, round(max(0.0, min(1.0, score)), 4), reason))
        out.sort(key=lambda x: -x[1])
        return out

    # -- qo'lda tasdiqlash ----------------------------------------------
    def confirm(self, line_id, product_id, raw_name=None):
        """Buxgalter tanlovini saqlaydi - bu juftlik boshqa so'ralmaydi."""
        cx = self.cx
        if raw_name is None:
            r = cx.execute("SELECT raw_name FROM doc_line WHERE id=?", (line_id,)).fetchone()
            raw_name = r["raw_name"] if r else ""
        DB.add_alias(cx, product_id, raw_name, source="manual")
        self.by_alias[C.norm_name(raw_name)] = product_id
        cx.execute(
            "UPDATE doc_line SET product_id=?, match_method='manual', match_score=1.0 "
            "WHERE id=?", (product_id, line_id))
        # bir xil nomli barcha bog'lanmagan satrlarni ham biriktirish
        n = C.norm_name(raw_name)
        cx.execute(
            "UPDATE doc_line SET product_id=?, match_method='manual', match_score=1.0 "
            "WHERE product_id IS NULL AND norm_name=?", (product_id, n))

    def unmatched_groups(self, year=None):
        """
        Bog'lanmagan sotuv satrlarini NOM bo'yicha guruhlab beradi.
        Buxgalter 597 satrni emas, 30-40 ta nomni ko'radi.
        """
        sql = ("SELECT l.norm_name, MIN(l.raw_name) raw_name, MIN(l.mxik) mxik, "
               "       MIN(l.barcode) barcode, COUNT(*) n, "
               "       SUM(CAST(l.qty AS REAL)) qty, "
               "       SUM(CAST(l.amount_gross AS REAL)) amount, "
               "       MIN(l.id) line_id "
               "FROM doc_line l JOIN document d ON d.id=l.document_id "
               "WHERE l.kind='chiqim' AND l.product_id IS NULL ")
        args = []
        if year is not None:
            sql += " AND d.doc_year=? "
            args.append(year)
        sql += " GROUP BY l.norm_name ORDER BY amount DESC"
        return self.cx.execute(sql, args).fetchall()

    def product_choices(self, query="", limit=200):
        """Qo'lda tanlash ro'yxati uchun mahsulotlar."""
        q = C.norm_name(query)
        rows = self.cx.execute(
            "SELECT p.id, p.canon_name, p.mxik, p.unit, "
            "  (SELECT COUNT(*) FROM stock_lot s WHERE s.product_id=p.id) lots, "
            "  (SELECT SUM(CAST(s.qty_left AS REAL)) FROM stock_lot s WHERE s.product_id=p.id) qty_left "
            "FROM product p ORDER BY p.canon_name").fetchall()
        if not q:
            return rows[:limit]
        # So'z o'zagi bo'yicha ham: "машины" -> "машина", "машинки" (rus tilidagi
        # qo'shimchalar buxgalterni chalg'itmasin)
        stems = [t[:max(4, len(t) - 2)] for t in q.split() if len(t) >= 3]
        out = []
        for r in rows:
            n = C.norm_name(r["canon_name"])
            if q in n:
                out.append((0, r))
            elif stems and all(st in n for st in stems):
                out.append((0.3, r))
            else:
                sc = difflib.SequenceMatcher(None, q, n).ratio()
                if sc > 0.45:
                    out.append((1 - sc, r))
        out.sort(key=lambda x: x[0])
        return [r for _, r in out[:limit]]


def auto_match_all(cx, engine=None, year=None, progress=None):
    """
    Barcha bog'lanmagan chiqim satrlarini avtomatik moslashtirishga urinadi.
    (avtomatik_bogʻlangan, qolgan) qaytaradi.
    """
    eng = engine or MatchEngine(cx)
    sql = ("SELECT l.* FROM doc_line l JOIN document d ON d.id=l.document_id "
           "WHERE l.kind='chiqim' AND l.product_id IS NULL")
    args = []
    if year is not None:
        sql += " AND d.doc_year=?"
        args.append(year)
    rows = cx.execute(sql, args).fetchall()

    done = 0
    total = len(rows)
    for i, r in enumerate(rows):
        line = dict(r)
        pid, meth, score, _sugg = eng.match(line)
        if pid and score >= AUTO_ACCEPT_SCORE:
            cx.execute(
                "UPDATE doc_line SET product_id=?, match_method=?, match_score=? WHERE id=?",
                (pid, meth, float(score), r["id"]))
            if meth in ("prefix", "fuzzy", "fuzzy+mxik"):
                DB.add_alias(cx, pid, r["raw_name"] or "", source="auto")
                eng.by_alias[C.norm_name(r["raw_name"] or "")] = pid
            done += 1
        if progress and (i % 50 == 0 or i == total - 1):
            progress(i + 1, total)
    return done, total - done
