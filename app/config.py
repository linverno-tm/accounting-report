# -*- coding: utf-8 -*-
"""
bh_config - o'zgarmas qiymatlar, ustun xaritalari, lug'atlar.

Bu modul hech narsani import qilmaydi (standart kutubxonadan tashqari),
shuning uchun exec tartibida birinchi turadi.
"""

import os
import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP

VERSION = "1.3.0"
APP_TITLE = "Buxgalteriya hisoboti generatori"
SCHEMA_VERSION = 1

# Chek fayli o'quvchisining versiyasi. Oshirilsa, bazadagi eski chek
# fayllari bir marta qayta o'qiladi (bh_fifo.reparse_old_checks).
#   2 - barcha varaqlar va <dimension> dan keyingi satrlar ham o'qiladi
CHECKS_PARSER_VER = 2

# ---------------------------------------------------------------------------
# Soliq / narx qoidalari
#
# Hisobotdagi "sotish narxi / tannarx" nisbatini sanab chiqqanda topilgani:
#   1.1536 = 1.03 x 1.12   (ustama 3%,  QQS 12%)   2023: 238, 2024: 392, 2025: 393 satr
#   1.2320 = 1.10 x 1.12   (ustama 10%, QQS 12%)   2024: 463, 2025: 222 satr
#   1.1760 = 1.05 x 1.12   (ustama 5%,  QQS 12%)   2026: 612 satr
#
# Ya'ni ustama yillik emas, satr bo'yicha. Shuning uchun:
#   default -> yillik stavka, ustidan mahsulot/ta'minotchi qoidasi, undan
#   ustidan satrning o'z qiymati.
# ---------------------------------------------------------------------------
VAT_RATE = Decimal("0.12")

DEFAULT_MARKUP_BY_YEAR = {
    2023: Decimal("0.03"),
    2024: Decimal("0.10"),
    2025: Decimal("0.03"),
    2026: Decimal("0.05"),
}
DEFAULT_MARKUP = Decimal("0.05")

MONEY_Q = Decimal("0.0001")   # pul: 4 kasr (hisobotdagi kabi)
QTY_Q = Decimal("0.0001")     # miqdor: 4 kasr


def money(x):
    """Istalgan qiymatni Decimal pulga aylantiradi. float ISHLATILMAYDI."""
    if x is None or x == "":
        return Decimal("0")
    if isinstance(x, Decimal):
        return x.quantize(MONEY_Q, rounding=ROUND_HALF_UP)
    return Decimal(str(x)).quantize(MONEY_Q, rounding=ROUND_HALF_UP)


def qty(x):
    if x is None or x == "":
        return Decimal("0")
    if isinstance(x, Decimal):
        return x.quantize(QTY_Q, rounding=ROUND_HALF_UP)
    return Decimal(str(x)).quantize(QTY_Q, rounding=ROUND_HALF_UP)


def vat_from_gross(gross):
    """QQS ichida bo'lgan summadan QQS ajratish: summa x 12 / 112."""
    g = money(gross)
    return money(g * VAT_RATE / (Decimal("1") + VAT_RATE))


def net_from_gross(gross):
    g = money(gross)
    return money(g / (Decimal("1") + VAT_RATE))


def sale_price_from_cost(cost, markup):
    """tannarx -> sotish narxi (tannarx + ustama + QQS)."""
    c = money(cost)
    m = Decimal(str(markup))
    return money(c * (Decimal("1") + m) * (Decimal("1") + VAT_RATE))


# ---------------------------------------------------------------------------
# Raqamlarni matndan ajratish
#
# Manbalarda: "14 906 785,71", "1 788 814.29", "50,000000", "\xa0" (nbsp)
# ---------------------------------------------------------------------------
_NUM_CLEAN = re.compile(r"[^\d,.\-]")


def parse_number(s):
    """Matndan Decimal. Ajratib bo'lmasa None."""
    if s is None:
        return None
    if isinstance(s, (int, float, Decimal)):
        return Decimal(str(s))
    t = str(s).strip()
    if not t:
        return None
    t = t.replace("\xa0", "").replace(" ", "").replace(" ", "")
    t = _NUM_CLEAN.sub("", t)
    if not t or t in ("-", ".", ","):
        return None
    # "1.234,56" (evropa) va "1,234.56" (ingliz) farqi
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        # vergul kasr ajratkichi deb qaraladi
        t = t.replace(",", ".")
    # ko'p nuqta bo'lsa oxirgisidan boshqasi minglik ajratkich
    if t.count(".") > 1:
        head, _, tail = t.rpartition(".")
        t = head.replace(".", "") + "." + tail
    try:
        return Decimal(t)
    except Exception:
        return None


_DATE_PATTERNS = [
    (re.compile(r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})"), (2, 1, 0)),  # 21.07.2026
    (re.compile(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})"), (0, 1, 2)),  # 2026-07-21
]


def parse_date(s):
    """Matndan datetime.date. Ajratib bo'lmasa None."""
    import datetime

    if s is None:
        return None
    if isinstance(s, datetime.datetime):
        return s.date()
    if isinstance(s, datetime.date):
        return s
    # Excel sana raqami (1900 tizimi): 46027.41 -> 2026-01-05.
    # Katak "Sana" formatisiz saqlangan bo'lsa openpyxl shunday qaytaradi.
    if isinstance(s, (int, float, Decimal)) and not isinstance(s, bool):
        if 20000 <= float(s) <= 80000:          # ~1954..2119 yillar
            return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(s)))
        return None
    t = str(s).strip()
    if re.match(r"^\d{5}(\.\d+)?$", t):
        return parse_date(float(t))
    if not t:
        return None
    for rx, order in _DATE_PATTERNS:
        m = rx.match(t)
        if m:
            try:
                g = m.groups()
                return datetime.date(
                    int(g[order[0]]), int(g[order[1]]), int(g[order[2]])
                )
            except ValueError:
                return None
    return None


# ---------------------------------------------------------------------------
# Nom normalizatsiyasi
#
# Kassa mahsulot nomini 63 BELGIDA KESADI (tekshirildi: касса 2026 da 410 satr,
# касса 2025 da 300 satr aynan 63 belgi). Shuning uchun moslashtirishda
# prefiks solishtirish ishlatiladi.
# ---------------------------------------------------------------------------
POS_NAME_MAX_LEN = 63

_PUNCT = re.compile(r"[^\w\s]", re.UNICODE)
_WS = re.compile(r"\s+")

# Kirill <-> lotin aralashib ketgan joylar (ko'rinishi bir xil harflar)
_LOOKALIKE = {
    "а": "a", "в": "b", "е": "e", "к": "k", "м": "m", "н": "h", "о": "o",
    "р": "p", "с": "c", "т": "t", "у": "y", "х": "x",
}


def norm_name(s):
    """Solishtirish uchun nomni normallashtirish."""
    if not s:
        return ""
    t = unicodedata.normalize("NFKC", str(s)).strip().lower()
    t = t.replace("ё", "е").replace("ʼ", "'").replace("’", "'").replace("`", "'")
    t = t.replace("№", "n").replace("&nbsp;", " ").replace("\xa0", " ")
    t = _PUNCT.sub(" ", t)
    t = _WS.sub(" ", t).strip()
    return t


def norm_name_aggressive(s):
    """Kirill/lotin ko'rinishdosh harflarni ham birlashtiradi."""
    t = norm_name(s)
    return "".join(_LOOKALIKE.get(ch, ch) for ch in t)


def name_prefix_key(s, n=POS_NAME_MAX_LEN):
    """
    Kassa kesgan uzunlikdagi prefiks kaliti.

    XOM nom avval kesiladi, keyin normallashtiriladi - chunki normalizatsiya
    tinish belgilarini olib tashlab uzunlikni o'zgartiradi. Agar avval
    normallashtirsak, kesilish nuqtasi siljib ketadi.
    """
    if not s:
        return ""
    return norm_name(str(s)[:n])


def is_truncated_match(pos_name, full_name):
    """
    Kassa nomi to'liq nomning kesilgan ko'rinishimi?

    Kassa 63 belgida kesadi, shuning uchun to'liq nomning normallashtirilgan
    ko'rinishi kassa nomining normallashtirilgan ko'rinishi bilan boshlanishi
    kerak.
    """
    p = norm_name(pos_name)
    f = norm_name(full_name)
    if not p or not f:
        return False
    if p == f:
        return True
    # Faqat haqiqatan kesilgan bo'lsa (xom nom chegaraga yetgan bo'lsa)
    if len(str(pos_name).strip()) < POS_NAME_MAX_LEN - 2:
        return False
    return f.startswith(p[: max(10, len(p) - 2)])


# ---------------------------------------------------------------------------
# O'lchov birliklari lug'ati
#
# Hisobotda bir narsa 3 xil yozilgan:
#   'шт.', 'dona', 'дона', 'шт. (потребительская коробка)',
#   'dona (istemol qutisi)', "dona (iste'mol qutisi)", 'потребительская
#   коробка=1 шта', ...
# ---------------------------------------------------------------------------
UNIT_CANON = {
    "dona": "dona",
    "донa": "dona",
    "дона": "dona",
    "шт": "dona",
    "шт.": "dona",
    "штука": "dona",
    "piece": "dona",
    "pcs": "dona",
    "kg": "kg",
    "килограмм": "kg",
    "кг": "kg",
    "kilogramm": "kg",
    "metr": "metr",
    "метр": "metr",
    "м": "metr",
    "kv. metr": "kv.metr",
    "кв. метр": "kv.metr",
    "kvmetr": "kv.metr",
    "litr": "litr",
    "литр": "litr",
    "komplekt": "komplekt",
    "комплект": "komplekt",
    "комплект/набор": "komplekt",
    "набор": "komplekt",
    "to'plam": "komplekt",
    "quti": "quti",
    "коробка": "quti",
    # "потребительская коробка=1 шта" kabi yozuvlar - aslida dona
    "потребительская коробка": "dona",
    "истемол кутиси": "dona",
    "istemol qutisi": "dona",
    "iste'mol qutisi": "dona",
    "упаковка": "qadoq",
    "qadoq": "qadoq",
    "пачка": "qadoq",
    "лист": "list",
    "list": "list",
}

# Qavs ichidagi izoh ("шт. (потребительская коробка)") asosiy birlikni
# o'zgartirmaydi - u qadoqlash turi.
_UNIT_PAREN = re.compile(r"\s*[\(\[].*?[\)\]]\s*")
_UNIT_EQ = re.compile(r"\s*=.*$")


def norm_unit(s):
    """O'lchov birligini yagona ko'rinishga keltiradi."""
    if not s:
        return ""
    t = unicodedata.normalize("NFKC", str(s)).strip().lower()
    t = t.replace("’", "'").replace("ʼ", "'")
    t = _UNIT_EQ.sub("", t)
    base = _UNIT_PAREN.sub(" ", t).strip().strip(".,;")
    base = _WS.sub(" ", base)
    if base in UNIT_CANON:
        return UNIT_CANON[base]
    nodot = base.rstrip(".")
    if nodot in UNIT_CANON:
        return UNIT_CANON[nodot]
    first = base.split(" ")[0].strip(".")
    if first in UNIT_CANON:
        return UNIT_CANON[first]
    return base or ""


def unit_packaging(s):
    """"шт. (потребительская коробка)" -> "потребительская коробка"."""
    if not s:
        return ""
    m = re.search(r"[\(\[](.*?)[\)\]]", str(s))
    return _WS.sub(" ", m.group(1)).strip() if m else ""


# ---------------------------------------------------------------------------
# "Маркировкаланган" ustuni
#
# Bitta ustunda aralash: 'маркировкаси', 'маркирокаланган' (xato yozilgan),
# 'без маркировкасиз', 'йўқ', 'йук', '-', '704.0', '35.0',
# '010478006301597021EI' (haqiqiy datamatrix).
# ---------------------------------------------------------------------------
_MARK_NO = {"", "-", "йўқ", "йук", "yoq", "yo'q", "без маркировкасиз",
            "без маркировки", "маркировкасиз", "нет", "no", "0"}
_MARK_YES = {"маркировкаси", "маркирокаланган", "маркировкаланган",
             "маркировка", "ha", "да", "bor", "yes", "1"}
_DATAMATRIX = re.compile(r"^01\d{14}[0-9A-Za-z]{6,}")


def parse_marking(v):
    """(is_marked: bool|None, marking_code: str) qaytaradi."""
    if v is None:
        return None, ""
    t = str(v).strip()
    if not t:
        return None, ""
    low = t.lower()
    if _DATAMATRIX.match(t):
        return True, t
    if low in _MARK_YES:
        return True, ""
    if low in _MARK_NO:
        return False, ""
    # '704.0', '35.0' kabi raqamlar - ma'nosiz, aniqlanmagan deb qaraladi
    if re.match(r"^\d+(\.\d+)?$", t):
        return None, ""
    if len(t) >= 20:
        return True, t
    return None, ""


# ---------------------------------------------------------------------------
# MANBA USTUNLARI
#
# FAKTURA fayllarida ustunlar soni O'ZGARIB TURADI (22 fayldan: 10 ustun -
# 11 ta, 11 ustun - 6 ta, 12 ustun - 4 ta). Shuning uchun HECH QACHON
# indeks bo'yicha o'qilmaydi, faqat sarlavha matni bo'yicha.
# ---------------------------------------------------------------------------
FAKTURA_HEADER_MAP = [
    ("row_no", ("№",)),
    ("name", ("маҳсулот номи", "махсулот номи", "mahsulot nomi", "наименование")),
    ("marking", ("маркировка коди", "markirovka kodi")),
    ("mxik", ("идентификация коди", "identifikatsiya kodi", "мхик", "икпу")),
    ("barcode", ("штрих коди", "штрих код", "shtrix kod")),
    ("unit", ("ўлчов бирлиги", "улчов бирлиги", "o'lchov birligi", "единица")),
    ("qty", ("миқдор", "микдор", "miqdor", "количество")),
    ("price", ("нарҳ", "нарх", "narx", "цена")),
    ("amount_net", ("етказиб бериш қиймати", "етказиб бериш киймати",
                    "yetkazib berish qiymati")),
    ("vat_rate", ("ставка",)),
    ("vat_amount", ("сумма",)),
    ("amount_gross", ("ҳисобга олган ҳолда қиймати", "хисобга олган холда киймати",
                      "ккс билан", "қҚс билан")),
    ("origin", ("келиб чиқиши", "kelib chiqishi")),
]

# CHIQIM (checks-info) - ustunlar barqaror, 20 ta.
#
# !!! ENG MUHIM: "Нархи" ustuni NARX EMAS, SATR SUMMASI (QQS ichida).
#     Chek-chek tekshirildi: satrlar yig'indisi = chekning "Жами карта" si.
#     Birlik narxi = Нархи / Миқдори
#
# !!! "Жами нақд пул" / "Жами банк карта" / "Жами ҚҚС" - CHEK darajasidagi
#     jami, lekin chekning HAR SATRIGA takrorlanadi. Satr bo'yicha yig'ish
#     tushumni 2.26 barobar shishirib yuboradi.
CHECKS_COLUMNS = [
    "product_id",      # 0  Махсулот Идси  ("...-0", "...-1" = chek ichidagi satr)
    "tin",             # 1  СТИР/ЖИШШР
    "pos_id",          # 2  ФМ рақами
    "check_dt",        # 3  Чек санаси
    "check_no",        # 4  Чек рақами
    "name",            # 5  Маҳсулот номи   (63 belgida kesiladi!)
    "qty",             # 6  Миқдори
    "line_gross",      # 7  Нархи  <-- SUMMA, narx emas
    "discount_sum",    # 8  Чегирма суммаси
    "bonus_sum",       # 9  Дисконт суммаси
    "line_vat",        # 10 ҚҚС суммаси    (= line_gross x 12/112)
    "check_cash",      # 11 Жами нақд пул      <-- chek jami, takrorlanadi
    "check_card",      # 12 Жами банк карта    <-- chek jami, takrorlanadi
    "check_vat",       # 13 Жами ҚҚС           <-- chek jami, takrorlanadi
    "mxik",            # 14 Маҳсулот коди
    "unit_code",       # 15 Ўлчов бирлиги коди
    "barcode",         # 16 Штрих код
    "agent_tin",       # 17 Воситачи СТИРи
    "marking",         # 18 Маркировка коди
    "check_type",      # 19 Чек тури  (Сотув / Қайтариш)
]

CHECK_TYPE_SALE = "сотув"
CHECK_TYPE_RETURN = ("қайтариш", "кайтариш", "qaytarish", "возврат")


def is_return_check(t):
    return norm_name(t) in {norm_name(x) for x in CHECK_TYPE_RETURN}


# ---------------------------------------------------------------------------
# HISOBOT (ТХ varag'i) ustunlari - mavjud faylning aynan tuzilishi
# ---------------------------------------------------------------------------
TX_COLUMNS = [
    (0,  "№",                        6),
    (1,  "санаси",                  11),
    (2,  "Товар номи",              46),
    (3,  "Маркировкаланган",        16),
    (4,  "МХИК Коди",               46),
    (5,  "Ўлчов бирлиги",           18),
    (6,  "1 бирликни сотиш нархи\n(таннарх+устама+ҚҚС)", 16),
    (7,  "Микдори",                 10),
    (8,  "Нархи",                   14),
    (9,  "суммаси",                 16),
    (10, "Микдори",                 10),
    (11, "Нархи",                   14),
    (12, "суммаси",                 16),
    (13, "Микдори",                 10),
    (14, "Нархи",                   14),
    (15, "суммаси",                 16),
    (16, "Микдори",                 10),
    (17, "Нархи",                   14),
    (18, "суммаси",                 16),
    (19, "Микдори",                 10),
    (20, "Нархи",                   14),
    (21, "суммаси",                 16),
]

TX_GROUPS = [
    (7,  9,  "давр бошига қолдиқ (таннарх)"),
    (10, 12, "Кирим (таннарх)"),
    (13, 15, "Чиқим (таннарх)"),
    (16, 18, "давр охирига қолдиқ таннархда"),
    (19, 21, "Чиқим сотиш нархида ҚҚС билан"),
]

KASSA_HEADERS = [
    "Махсулот\nИдси", "СТИР/\nЖИШШР", "ФМ рақами", "Чек санаси", "Чек рақами",
    "Маҳсулот номи", "Миқдори", "Нархи", "Чегирма суммаси", "Дисконт суммаси",
    "ҚҚС суммаси", "Жами нақд пул", "Жами банк карта", "Жами ҚҚС",
    "Маҳсулот коди", "Ўлчов бирлиги коди", "Штрих код",
    "Воситачи СТИРи (ЖИШШРи)", "Маркировка\nкоди", "Чек тури",
]


# ---------------------------------------------------------------------------
# Tekshiruv (validatsiya) darajalari
# ---------------------------------------------------------------------------
SEVERITY_ERROR = "xato"
SEVERITY_WARN = "ogohlantirish"
SEVERITY_INFO = "ma'lumot"

ISSUE_TITLES = {
    "duplicate_doc": "Hujjat takrorlangan",
    "check_total_mismatch": "Chek jami satrlarga to'g'ri kelmadi",
    "vat_mismatch": "QQS summasi 12% ga to'g'ri kelmadi",
    "negative_qty": "Manfiy miqdor",
    "negative_cost": "Manfiy tannarx",
    "no_stock": "Omborda yetarli tovar yo'q (kirim hujjati topilmadi)",
    "late_receipt": "Kirim hujjati sotuvdan keyingi sana bilan kiritilgan",
    "unmatched_sale": "Sotuv mahsuloti kirimga bog'lanmadi",
    "wrong_year": "Sana hisobot davriga tushmaydi",
    "bad_date": "Sana o'qilmadi",
    "unit_mismatch": "Kirim va chiqim o'lchov birligi har xil",
    "missing_price": "Narx yo'q yoki nol",
}


# ---------------------------------------------------------------------------
# Yo'llar
# ---------------------------------------------------------------------------
def app_data_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "BuxgalterHisobot")
    os.makedirs(d, exist_ok=True)
    return d


def db_path():
    return os.path.join(app_data_dir(), "hisobot.db")


def settings_path():
    return os.path.join(app_data_dir(), "settings.json")
