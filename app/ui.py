# -*- coding: utf-8 -*-
"""
bh_ui - tkinter interfeysi.

LAYOUT QOIDALARI (talab bo'yicha):

  1. Pastki panel HECH QACHON yo'qolmaydi.
     tkinter'da `pack` tartibi muhim: pastki panellar BIRINCHI pack qilinadi,
     asosiy maydon oxirida `expand=True` bilan. Shunda oyna qisqarganda
     asosiy maydon kichrayadi, tugmalar esa joyida qoladi.

  2. Tugmalar torayganda pastga O'TADI, kesilmaydi.
     FlowBar har `<Configure>` da qayta joylashtiradi: sig'masa yangi
     qatorga tushadi va panel balandligi o'sadi.

  3. Oyna ekranga qarab ochiladi.
     Ideal o'lcham ekranning 88% idan katta bo'lsa, kichraytiriladi.
     minsize hech qachon ekrandan katta bo'lmaydi.

  4. Matn kesilmaydi.
     Uzun yozuvlar `wraplength` bilan o'raladi, jadvallarda gorizontal
     aylantirgich bor, forma maydonlari vertikal aylantiriladi.

  5. DPI.
     Windows'da 125%/150% masshtabda matn xiralashmasligi va kesilmasligi
     uchun process DPI-aware qilinadi va shrift masshtablanadi.
"""

import os
import re
import sys
import queue
import threading
import traceback
import datetime
import webbrowser

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import bh_config as C
import bh_db as DB
import bh_parsers as P
import bh_matching as M
import bh_fifo as F
import bh_report as R


# ===========================================================================
# DPI va tema
# ===========================================================================
def enable_dpi_awareness():
    """Windows'da xira/kesilgan matnning oldini oladi."""
    if not sys.platform.startswith("win"):
        return 1.0
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # per-monitor
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)     # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return max(1.0, dpi / 96.0)
    except Exception:
        return 1.0


PALETTE = {
    "bg": "#f4f6f9",
    "card": "#ffffff",
    "ink": "#1c2434",
    "muted": "#5b6678",
    "line": "#d7dde7",
    "accent": "#1f6feb",
    "accent_dark": "#1a5ccc",
    "ok": "#1a7f4b",
    "warn": "#9a6700",
    "err": "#b42318",
    "warn_bg": "#fff6e0",
    "err_bg": "#fdeceb",
    "ok_bg": "#e8f5ee",
    "green": "#1a7f4b",
    "green_dark": "#146b3f",
    "grey_bg": "#eef0f4",
}


def setup_style(root, scale):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    # Buxgalter yoshi katta - harflar odatdagidan kattaroq (egasi talabi, 1.4.0)
    base = max(11, int(round(12 * scale)))
    # "*Font" option ttk yozuvlarining uslubdagi shriftini bosib ketardi (sarlavhalar
    # kattalashmasdi) - shuning uchun nomli standart shriftlar o'zgartiriladi.
    from tkinter import font as tkfont
    for nm in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
        try:
            tkfont.nametofont(nm).configure(family="Segoe UI", size=base)
        except tk.TclError:
            pass

    f_base = ("Segoe UI", base)
    f_bold = ("Segoe UI", base, "bold")
    f_h1 = ("Segoe UI", base + 6, "bold")
    f_h2 = ("Segoe UI", base + 3, "bold")
    f_small = ("Segoe UI", max(10, base - 1))
    f_mono = ("Consolas", max(8, base - 1))

    root.configure(bg=PALETTE["bg"])
    style.configure(".", font=f_base, background=PALETTE["bg"],
                    foreground=PALETTE["ink"])
    style.configure("TFrame", background=PALETTE["bg"])
    # Card.TFrame - tashqi ramka (chegarali).
    # CardIn.TFrame - ichki konteyner: ramkaSIZ, aks holda ikkilangan chiziq
    # chiqadi va kartochka ichida keraksiz bo'linish ko'rinadi.
    style.configure("Card.TFrame", background=PALETTE["card"],
                    relief="solid", borderwidth=1)
    style.configure("CardIn.TFrame", background=PALETTE["card"],
                    relief="flat", borderwidth=0)
    style.configure("Drop.TFrame", background="#eef3fb",
                    relief="solid", borderwidth=1)
    style.configure("DropIn.TFrame", background="#eef3fb",
                    relief="flat", borderwidth=0)
    style.configure("TLabel", background=PALETTE["bg"], foreground=PALETTE["ink"])
    style.configure("Card.TLabel", background=PALETTE["card"])
    style.configure("H1.TLabel", font=f_h1, background=PALETTE["bg"])
    style.configure("H2.TLabel", font=f_h2, background=PALETTE["card"])
    style.configure("Muted.TLabel", foreground=PALETTE["muted"], font=f_small,
                    background=PALETTE["bg"])
    style.configure("CardMuted.TLabel", foreground=PALETTE["muted"], font=f_small,
                    background=PALETTE["card"])
    style.configure("Big.TLabel", font=("Segoe UI", base + 7, "bold"),
                    background=PALETTE["card"])
    style.configure("Ok.TLabel", foreground=PALETTE["ok"], background=PALETTE["card"])
    style.configure("Warn.TLabel", foreground=PALETTE["warn"], background=PALETTE["card"])
    style.configure("Err.TLabel", foreground=PALETTE["err"], background=PALETTE["card"])

    pad = (int(14 * scale), int(7 * scale))
    style.configure("TButton", font=f_base, padding=pad)
    style.configure("Accent.TButton", font=f_bold, padding=pad,
                    background=PALETTE["accent"], foreground="#ffffff",
                    borderwidth=0)
    style.map("Accent.TButton",
              background=[("active", PALETTE["accent_dark"]),
                          ("disabled", "#9fb8e8")])
    style.configure("Ghost.TButton", font=f_base, padding=pad)
    # Katta tugmalar: asosiy amallar (fayl qo'shish, hisobot)
    bigpad = (int(22 * scale), int(12 * scale))
    style.configure("Big.TButton", font=("Segoe UI", base + 2, "bold"), padding=bigpad,
                    background=PALETTE["accent"], foreground="#ffffff", borderwidth=0)
    style.map("Big.TButton", background=[("active", PALETTE["accent_dark"]),
                                         ("disabled", "#9fb8e8")])
    style.configure("BigGhost.TButton", font=("Segoe UI", base + 2), padding=bigpad)
    style.configure("Green.TButton", font=("Segoe UI", base + 4, "bold"),
                    padding=(int(28 * scale), int(16 * scale)),
                    background=PALETTE["green"], foreground="#ffffff", borderwidth=0)
    style.map("Green.TButton", background=[("active", PALETTE["green_dark"]),
                                           ("disabled", "#9cc9b0")])
    style.configure("Step.TLabel", font=("Segoe UI", base + 4, "bold"),
                    background=PALETTE["card"], foreground=PALETTE["ink"])
    style.configure("Text.TLabel", font=f_base, background=PALETTE["card"],
                    foreground=PALETTE["ink"])
    style.configure("Result.TLabel", font=("Segoe UI", base + 1), background=PALETTE["card"],
                    foreground=PALETTE["ink"])
    style.configure("TRadiobutton", background=PALETTE["card"], font=("Segoe UI", base + 1))
    style.configure("TCheckbutton", background=PALETTE["bg"], font=f_base)

    style.configure("TNotebook", background=PALETTE["bg"], borderwidth=0)
    style.configure("TNotebook.Tab", font=("Segoe UI", base + 1),
                    padding=(int(18 * scale), int(10 * scale)))
    style.map("TNotebook.Tab",
              background=[("selected", PALETTE["card"])],
              foreground=[("selected", PALETTE["accent"])])

    rowh = int(round(30 * scale))
    style.configure("Treeview", font=f_base, rowheight=rowh,
                    background=PALETTE["card"], fieldbackground=PALETTE["card"],
                    borderwidth=1)
    style.configure("Treeview.Heading", font=f_bold,
                    padding=(int(6 * scale), int(5 * scale)))
    style.map("Treeview", background=[("selected", PALETTE["accent"])],
              foreground=[("selected", "#ffffff")])

    style.configure("TEntry", padding=int(5 * scale))
    style.configure("TCombobox", padding=int(5 * scale))
    style.configure("Horizontal.TProgressbar", thickness=int(9 * scale),
                    background=PALETTE["accent"])

    return {"base": f_base, "bold": f_bold, "h1": f_h1, "h2": f_h2,
            "small": f_small, "mono": f_mono, "scale": scale}


# ===========================================================================
# Javob beruvchan vidjetlar
# ===========================================================================
class FlowBar(ttk.Frame):
    """
    Tugmalar paneli. Oyna torayganda tugmalar PASTGA O'TADI -
    hech qachon kesilmaydi va ekran tashqarisiga chiqib ketmaydi.

    left  guruh: chapdan boshlab oqadi
    right guruh: o'ngga tiralib turadi, sig'masa yangi qatorga tushadi
    """

    def __init__(self, master, scale=1.0, **kw):
        super().__init__(master, **kw)
        self.gap = max(6, int(8 * scale))
        self.vgap = max(5, int(6 * scale))
        self._left = []
        self._right = []
        self._last_w = -1
        self.bind("<Configure>", self._on_configure)

    def add(self, widget, side="left"):
        (self._left if side == "left" else self._right).append(widget)
        widget.place(x=-4000, y=0)          # o'lchanguncha ko'rinmasin
        self.after_idle(self._relayout)
        return widget

    def _on_configure(self, event):
        if event.width != self._last_w:
            self._last_w = event.width
            self._relayout()

    def _relayout(self):
        width = self.winfo_width()
        if width <= 1:
            self.after(30, self._relayout)
            return

        items = [(w, w.winfo_reqwidth(), w.winfo_reqheight())
                 for w in self._left + self._right if w.winfo_exists()]
        if not items:
            return
        rowh = max(h for _w, _rw, h in items)

        # Qatorlarga bo'lish
        rows, cur, cur_w = [], [], 0
        for w, rw, _h in items:
            need = rw if not cur else rw + self.gap
            if cur and cur_w + need > width - self.gap:
                rows.append((cur, cur_w))
                cur, cur_w = [(w, rw)], rw
            else:
                cur.append((w, rw))
                cur_w += need
        if cur:
            rows.append((cur, cur_w))

        nleft = len(self._left)
        y = self.vgap
        placed = 0
        for ri, (row, roww) in enumerate(rows):
            # Oxirgi qatorda faqat o'ng guruh qolgan bo'lsa - o'ngga tirash
            only_right = all(placed + i >= nleft for i in range(len(row)))
            x = (width - roww - self.gap) if (only_right and roww + 2 * self.gap <= width) \
                else self.gap
            for w, rw in row:
                w.place(x=int(x), y=int(y), height=rowh)
                x += rw + self.gap
                placed += 1
            y += rowh + self.vgap

        self.configure(height=int(y))
        self.pack_propagate(False)
        self.grid_propagate(False)


class ScrollFrame(ttk.Frame):
    """
    Vertikal aylantiriladigan konteyner. Ichidagi forma oynaga sig'masa
    matn kesilmaydi - aylantirgich paydo bo'ladi.
    """

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0,
                                background=PALETTE["bg"])
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._on_scroll)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.body, anchor="nw")

        self.body.bind("<Configure>", self._on_body)
        self.canvas.bind("<Configure>", self._on_canvas)
        self.canvas.bind("<Enter>", lambda e: self._bind_wheel(True))
        self.canvas.bind("<Leave>", lambda e: self._bind_wheel(False))

    def _on_scroll(self, lo, hi):
        # Aylantirgich faqat kerak bo'lganda ko'rinadi
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.vbar.pack_forget()
        else:
            self.vbar.pack(side="right", fill="y")
        self.vbar.set(lo, hi)

    def _on_body(self, _e=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas(self, e):
        self.canvas.itemconfigure(self._win, width=e.width)

    def _bind_wheel(self, on):
        if on:
            self.canvas.bind_all("<MouseWheel>", self._wheel)
        else:
            self.canvas.unbind_all("<MouseWheel>")

    def _wheel(self, e):
        first, last = self.canvas.yview()
        if first <= 0.0 and last >= 1.0:
            return
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")


class Table(ttk.Frame):
    """
    Treeview + ikkala aylantirgich.

    Ustunlar: (kalit, sarlavha, eng_kichik_kenglik, cho'ziladimi, tekislash)
    Cho'ziladigan ustun oyna kengayganda o'sadi, qolganlari eng kichik
    kengligidan pastga tushmaydi - shuning uchun raqamlar kesilmaydi.
    """

    def __init__(self, master, columns, scale=1.0, height=12, on_select=None,
                 on_double=None, **kw):
        super().__init__(master, **kw)
        self.cols = columns
        keys = [c[0] for c in columns]

        self.tree = ttk.Treeview(self, columns=keys, show="headings",
                                 height=height, selectmode="browse")
        vb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vb.set, xscrollcommand=hb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vb.grid(row=0, column=1, sticky="ns")
        hb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        for key, title, minw, stretch, anchor in columns:
            w = int(minw * scale)
            self.tree.heading(key, text=title,
                              command=lambda k=key: self._sort(k))
            self.tree.column(key, width=w, minwidth=w, stretch=bool(stretch),
                             anchor=anchor)

        self.tree.tag_configure("err", background=PALETTE["err_bg"])
        self.tree.tag_configure("warn", background=PALETTE["warn_bg"])
        self.tree.tag_configure("ok", background=PALETTE["ok_bg"])
        self.tree.tag_configure("odd", background="#fafbfd")

        self._sort_key = None
        self._sort_rev = False
        if on_select:
            self.tree.bind("<<TreeviewSelect>>", on_select)
        if on_double:
            self.tree.bind("<Double-1>", on_double)

    def clear(self):
        self.tree.delete(*self.tree.get_children())

    def add(self, values, tags=(), iid=None):
        return self.tree.insert("", "end", iid=iid, values=values, tags=tags)

    def fill(self, rows, tag_fn=None):
        self.clear()
        for i, r in enumerate(rows):
            tags = list(tag_fn(r)) if tag_fn else []
            if i % 2:
                tags.append("odd")
            self.tree.insert("", "end", iid=str(i), values=r, tags=tags)

    def selected_index(self):
        s = self.tree.selection()
        return int(s[0]) if s and s[0].isdigit() else None

    def _sort(self, key):
        idx = [c[0] for c in self.cols].index(key)
        rev = not self._sort_rev if self._sort_key == key else False
        self._sort_key, self._sort_rev = key, rev
        items = [(self.tree.set(i, key), i) for i in self.tree.get_children("")]

        def conv(v):
            t = str(v).replace(" ", "").replace(",", ".")
            try:
                return (0, float(t))
            except ValueError:
                return (1, str(v).lower())

        items.sort(key=lambda x: conv(x[0]), reverse=rev)
        for pos, (_v, iid) in enumerate(items):
            self.tree.move(iid, "", pos)


def card(master, title=None, fonts=None):
    """Oq fonli bo'lim. Ichki freym ataylab ramkasiz - ikkilangan chiziq chiqmasin."""
    outer = ttk.Frame(master, style="Card.TFrame")
    inner = ttk.Frame(outer, style="CardIn.TFrame")
    inner.pack(fill="both", expand=True, padx=12, pady=10)
    if title:
        ttk.Label(inner, text=title, style="H2.TLabel").pack(anchor="w", pady=(0, 8))
    return outer, inner


def wrapping_label(master, text, style="CardMuted.TLabel", pad=24):
    """Konteyner kengligiga qarab o'raladigan yozuv - matn kesilmaydi."""
    lbl = ttk.Label(master, text=text, style=style, justify="left", anchor="w")

    def resize(e):
        w = max(160, e.width - pad)
        if lbl.cget("wraplength") != w:
            lbl.configure(wraplength=w)

    master.bind("<Configure>", resize, add="+")
    return lbl


def fmt_money(v):
    try:
        return "{:,.2f}".format(float(v)).replace(",", " ")
    except (TypeError, ValueError):
        return ""


def fmt_qty(v):
    try:
        f = float(v)
        return "{:,.0f}".format(f).replace(",", " ") if f == int(f) \
            else "{:,.4f}".format(f).replace(",", " ")
    except (TypeError, ValueError):
        return ""


# ===========================================================================
# Explorer'dan sudrab tashlash (drag & drop)
# ===========================================================================
def enable_file_drop(widget, callback):
    """
    Vidjetga Explorer'dan fayl tashlash imkonini beradi.

    Kutubxona topilmasa jimgina o'chib qoladi - tugmalar orqali qo'shish
    baribir ishlayveradi. Shuning uchun dastur hech qachon shu sababdan
    ishlamay qolmaydi.

    Ishlatilgan backend nomini yoki None qaytaradi.
    """
    def deliver(paths):
        out = []
        for p in paths:
            if isinstance(p, bytes):
                for enc in ("utf-8", "cp1251", "mbcs"):
                    try:
                        p = p.decode(enc)
                        break
                    except (UnicodeDecodeError, LookupError):
                        continue
                else:
                    continue
            p = str(p).strip().strip("{}")
            if p:
                out.append(os.path.normpath(p))
        if out:
            try:
                callback(out)
            except Exception:
                traceback.print_exc()

    try:
        import windnd
        windnd.hook_dropfiles(widget, func=deliver)
        return "windnd"
    except Exception:
        pass

    try:
        from tkinterdnd2 import DND_FILES
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>",
                        lambda e: deliver(widget.tk.splitlist(e.data)))
        return "tkinterdnd2"
    except Exception:
        pass

    return None


# ===========================================================================
# Ish stoli yorlig'i ("o'rnatish")
# ===========================================================================
def app_exe_path():
    """
    Yorliq ko'rsatadigan fayl.

    .exe sifatida ishlayotgan bo'lsa - o'sha .exe. Ishlab chiqish rejimida
    (oddiy python) yorliq yasashning ma'nosi yo'q - None qaytadi.
    """
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return None


def _ps_quote(s):
    return "'" + str(s).replace("'", "''") + "'"


def create_shortcut(target, link_path, description="", icon=None):
    """
    Windows yorlig'ini (.lnk) yaratadi.

    pywin32 kerak emas - PowerShell'ning WScript.Shell obyekti ishlatiladi,
    shuning uchun .exe ichiga qo'shimcha kutubxona o'ralmaydi.
    """
    import subprocess

    os.makedirs(os.path.dirname(link_path), exist_ok=True)
    ps = (
        "$s = New-Object -ComObject WScript.Shell; "
        "$l = $s.CreateShortcut(%s); "
        "$l.TargetPath = %s; "
        "$l.WorkingDirectory = %s; "
        "$l.Description = %s; "
        % (_ps_quote(link_path), _ps_quote(target),
           _ps_quote(os.path.dirname(target)), _ps_quote(description))
    )
    if icon:
        ps += "$l.IconLocation = %s; " % _ps_quote(icon)
    ps += "$l.Save()"

    flags = 0x08000000 if sys.platform.startswith("win") else 0   # oyna chiqmasin
    r = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive",
         "-ExecutionPolicy", "Bypass", "-Command", ps],
        capture_output=True, text=True, creationflags=flags, timeout=30)
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout or "").strip()[:400])
    return link_path


def desktop_dir():
    """Ish stoli papkasi. OneDrive'ga ko'chirilgan bo'lsa ham topadi."""
    import subprocess

    try:
        flags = 0x08000000 if sys.platform.startswith("win") else 0
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "[Environment]::GetFolderPath('Desktop')"],
            capture_output=True, text=True, creationflags=flags, timeout=20)
        p = (r.stdout or "").strip()
        if p and os.path.isdir(p):
            return p
    except Exception:
        pass
    for cand in (os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop"),
                 os.path.join(os.path.expanduser("~"), "Desktop")):
        if os.path.isdir(cand):
            return cand
    return os.path.expanduser("~")


def install_shortcuts(title=None):
    """
    Ish stoliga va Boshlash menyusiga yorliq qo'yadi.

    Yaratilgan yo'llar ro'yxatini qaytaradi. .exe emas bo'lsa - bo'sh ro'yxat.
    """
    exe = app_exe_path()
    if not exe or not sys.platform.startswith("win"):
        return []

    title = title or C.APP_TITLE
    made = []
    targets = [
        os.path.join(desktop_dir(), "%s.lnk" % title),
        os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")),
                     "Microsoft", "Windows", "Start Menu", "Programs",
                     "%s.lnk" % title),
    ]
    for link in targets:
        try:
            create_shortcut(exe, link,
                            description="Faktura va kassa cheklaridan hisobot",
                            icon="%s,0" % exe)
            made.append(link)
        except Exception:
            continue
    return made


# ===========================================================================
# Bitta nusxa ("o'rnatish")
#
# .exe ni buxgalter qayerdan ochsa o'sha joydan ishlardi: yangisini yuklab olsa
# kompyuterda ikkinchi nusxa paydo bo'lardi ("BuxgalterHisobot (1).exe",
# Telegram papkasida yana bittasi...). Egasining talabi: "bitta ilova ikkita
# bo'lib qolmasin".
#
# Endi doimiy joy bor: %LOCALAPPDATA%\BuxgalterHisobot\BuxgalterHisobot.exe.
#   1. Boshqa joydan ochilsa - o'zini shu yerga ko'chiradi (eski versiya
#      ustidan yoziladi), yorliqlarni shu yerga qaratadi, o'sha nusxani
#      ishga tushirib o'zi yopiladi va yopilgandan keyin o'chiriladi.
#   2. Doimiy joydan ochilganda - odatiy papkalardagi boshqa nusxalarni
#      o'chiradi.
# O'chiriladigan fayl: nomi BuxgalterHisobot*.exe, hajmi 5 MB dan katta
# (PyInstaller to'plami) va doimiy nusxaning o'zi emas. Ishlab turgan nusxa
# qulflangan bo'ladi - o'chmaydi, keyingi ochilishda qayta urinadi.
# ===========================================================================
INSTALL_EXE_NAME = "BuxgalterHisobot.exe"
_COPY_NAME = re.compile(r"^buxgalterhisobot(\s*\(\d+\)|[\s_-]*v?[\d.]+)?\.exe$", re.IGNORECASE)
_MIN_EXE_BYTES = 5 * 1024 * 1024


def install_path():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "BuxgalterHisobot", INSTALL_EXE_NAME)


def _same_file(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _stray_dirs():
    home = os.path.expanduser("~")
    dirs = [desktop_dir(), os.path.join(home, "Desktop"), os.path.join(home, "OneDrive", "Desktop"),
            os.path.join(home, "Downloads"), os.path.join(home, "Documents"),
            os.path.join(home, "Downloads", "Telegram Desktop"),
            os.path.join(home, "Documents", "Telegram Desktop")]
    out = []
    for d in dirs:
        if d and os.path.isdir(d) and not any(_same_file(d, x) for x in out):
            out.append(d)
    return out


def _is_app_copy(path):
    try:
        return (_COPY_NAME.match(os.path.basename(path)) is not None
                and os.path.isfile(path) and os.path.getsize(path) > _MIN_EXE_BYTES)
    except OSError:
        return False


def remove_stray_copies(keep):
    """Odatiy papkalardagi boshqa nusxalarni o'chiradi. O'chirilganlar ro'yxati."""
    removed = []
    for d in _stray_dirs():
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for n in names:
            p = os.path.join(d, n)
            if _same_file(p, keep) or not _is_app_copy(p):
                continue
            try:
                os.remove(p)
                removed.append(p)
            except OSError:
                pass          # ochiq turgan nusxa - keyingi safar
    return removed


def _remove_later(path, tries=40):
    """
    Eski (yuklab olingan) nusxani o'chiradi: u hali yopilayotgan bo'ladi, shuning uchun
    bir necha soniya qayta urinadi. Fon oqimida - oyna ham, tashqi buyruq ham yo'q
    (ilgari "cmd /c ping ... del" ishlatilardi - Windows 11 da qora oyna chiqardi).
    """
    import time

    def run():
        for _ in range(tries):
            try:
                if not os.path.exists(path):
                    return
                os.remove(path)
                return
            except OSError:
                time.sleep(0.5)
    threading.Thread(target=run, daemon=True).start()


def ensure_single_install():
    """
    True qaytarsa - doimiy nusxa ishga tushirildi, bu jarayon darhol yopilsin.
    Har qanday xatoda jim o'tadi: dastur baribir ochilishi kerak.
    """
    exe = app_exe_path()
    if not exe or not sys.platform.startswith("win"):
        return False
    target = install_path()
    try:
        if _same_file(exe, target):
            # Yangi o'rnatilgan nusxa: o'zini ochgan eski faylni o'chiradi
            if "--eski" in sys.argv:
                i = sys.argv.index("--eski")
                old = sys.argv[i + 1] if i + 1 < len(sys.argv) else ""
                if old and not _same_file(old, target) and _is_app_copy(old):
                    _remove_later(old)
            remove_stray_copies(target)
            return False
        # Boshqa joydan ochildi: o'zini doimiy joyga ko'chiradi. Eski nusxa
        # ochiq bo'lsa (qulflangan) - ko'chira olmaydi, shu joydan ishlayveradi.
        import shutil
        import subprocess
        os.makedirs(os.path.dirname(target), exist_ok=True)
        tmp = target + ".new"
        shutil.copy2(exe, tmp)
        try:
            os.replace(tmp, target)
        except OSError:
            try:
                os.remove(tmp)
            except OSError:
                pass
            return False
        try:
            install_shortcuts()
        except Exception:
            pass
        args = [target]
        if _is_app_copy(exe):
            args += ["--eski", exe]      # yangi nusxa shu faylni o'chiradi
        subprocess.Popen(args, cwd=os.path.dirname(target), close_fds=True,
                         creationflags=0x00000008)       # DETACHED_PROCESS (GUI, oyna yo'q)
        return True
    except Exception:
        return False


def expand_inputs(paths):
    """
    Tashlangan yoki tanlangan narsalarni fayl ro'yxatiga aylantiradi.

    Bitta fayl ham, 10 ta fayl ham, papka ham, aralash ham bo'lishi mumkin -
    hammasi bir xil ishlanadi.
    """
    files = []
    for p in paths:
        if os.path.isdir(p):
            files.extend(P.scan_folder(p))
        elif os.path.isfile(p):
            if p.lower().endswith(P.SUPPORTED_EXT) and \
                    not os.path.basename(p).startswith("~$"):
                files.append(p)
    seen = set()
    out = []
    for f in files:
        k = os.path.normcase(os.path.abspath(f))
        if k not in seen:
            seen.add(k)
            out.append(f)
    return out


# ===========================================================================
# Fon vazifasi (UI muzlab qolmasligi uchun)
# ===========================================================================
class Worker:
    """
    Uzoq ishlarni alohida oqimda bajaradi.

    SQLite ulanishi oqimlar orasida bo'lishilmaydi - ishchi oqim o'z
    ulanishini ochadi. UI oqimi natijani navbat orqali oladi.
    """

    def __init__(self, root, on_progress, on_done, on_error):
        self.root = root
        self.q = queue.Queue()
        self.on_progress = on_progress
        self.on_done = on_done
        self.on_error = on_error
        self.busy = False
        self._poll()

    def start(self, label, fn):
        if self.busy:
            return False
        self.busy = True
        self.q.put(("start", label, None))

        def progress(cur, total, msg=""):
            self.q.put(("progress", (cur, total, msg), None))

        def run():
            cx = None
            try:
                cx = DB.connect()
                result = fn(cx, progress)
                self.q.put(("done", label, result))
            except Exception:
                self.q.put(("error", label, traceback.format_exc()))
            finally:
                if cx is not None:
                    try:
                        cx.close()
                    except Exception:
                        pass

        threading.Thread(target=run, daemon=True).start()
        return True

    def _poll(self):
        try:
            while True:
                kind, a, b = self.q.get_nowait()
                if kind == "start":
                    self.on_progress(0, 0, a)
                elif kind == "progress":
                    self.on_progress(a[0], a[1], a[2])
                elif kind == "done":
                    self.busy = False
                    self.on_done(a, b)
                elif kind == "error":
                    self.busy = False
                    self.on_error(a, b)
        except queue.Empty:
            pass
        self.root.after(80, self._poll)


# ===========================================================================
# Asosiy oyna
# ===========================================================================
class App:
    def __init__(self, root, launcher_info=None):
        self.root = root
        self.info = launcher_info or {}
        self.scale = enable_dpi_awareness()
        self.fonts = setup_style(root, self.scale)
        self.cx = DB.connect()
        self.engine = M.MatchEngine(self.cx)
        self.queued_files = []
        self._unmatched = []
        self._candidates = []
        self._products = []
        self._last_report = None
        self._sha_cache = {}      # (yo'l, hajm, vaqt) -> (sha256, format)
        self._stock_dirty = False  # qo'lda bog'landi, ombor qayta hisoblansin
        self._stale = set()        # yangilanishi kerak bo'lgan og'ir bo'limlar

        root.title("%s - %s" % (C.APP_TITLE, C.VERSION))
        self._set_window_icon()
        self._setup_geometry()
        self._build()
        self.worker = Worker(root, self._on_progress, self._on_done, self._on_error)
        self.refresh_all()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        # Papkalar oldin ko'rsatilgan bo'lsa - darrov skanerlaymiz, buxgalter
        # har safar qo'lda bosmasin.
        root.after(120, self._startup_scan)
        root.after(400, self._first_run_setup)

    def _first_run_setup(self):
        """
        Birinchi ochilishda ish stoliga yorliq qo'yadi.

        Buxgalter .exe ni qayerga qo'yganini eslab o'tirmasin - yorliq
        ko'rinib tursin. Faqat BIR MARTA bajariladi; foydalanuvchi yorliqni
        o'chirsa qayta tiklanmaydi.
        """
        if DB.get_setting(self.cx, "shortcuts_installed"):
            return
        try:
            made = install_shortcuts()
        except Exception:
            made = []
        DB.set_setting(self.cx, "shortcuts_installed", True)
        if made:
            self.set_status("Ish stoliga yorliq qo'yildi")

    def _startup_scan(self):
        if (self.var_fak.get().strip() or self.var_chq.get().strip()):
            try:
                self.scan_folders()
            except Exception:
                pass

    def _set_window_icon(self):
        """
        Oyna sarlavhasi va vazifalar panelidagi belgi.

        .exe ga --icon bilan o'rnatilgan belgi jarayonga tegishli, lekin
        tkinter oynasi o'zining Tk patini ko'rsatib qolishi mumkin.
        Shuning uchun .ico ni aniq belgilaymiz. Topilmasa - jimgina
        o'tkazib yuboriladi (dastur bundan to'xtamaydi).
        """
        for d in (getattr(sys, "_MEIPASS", None),
                  os.path.dirname(os.path.abspath(sys.executable)),
                  os.path.dirname(os.path.dirname(os.path.abspath(__file__)))):
            if not d:
                continue
            p = os.path.join(d, "icon.ico")
            try:
                if os.path.isfile(p):
                    self.root.iconbitmap(default=p)
                    return p
            except Exception:
                continue
        return None

    # -- 3-qoida: oyna ekranga sig'adigan qilib ochiladi -----------------
    def _setup_geometry(self):
        root = self.root
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()

        ideal_w, ideal_h = int(1240 * self.scale), int(780 * self.scale)
        w = min(ideal_w, int(sw * 0.92))
        h = min(ideal_h, int(sh * 0.90))

        # Eng kichik o'lcham ham ekrandan oshmasin - aks holda oynani
        # umuman kichraytirib bo'lmay qoladi.
        #
        # Bu qiymat ataylab past: torayganda tugmalar FlowBar bilan pastga
        # o'tadi, formalar ScrollFrame bilan aylanadi, jadvallarda gorizontal
        # aylantirgich bor. Shuning uchun kichik oynada ham hech narsa
        # kesilmaydi - foydalanuvchini katta oynaga majburlash shart emas.
        min_w = min(int(620 * self.scale), int(sw * 0.90))
        min_h = min(int(460 * self.scale), int(sh * 0.85))
        root.minsize(min_w, min_h)

        saved = DB.get_setting(self.cx, "window_geometry")
        if saved and isinstance(saved, str) and "x" in saved:
            try:
                gw, gh = (int(x) for x in saved.split("+")[0].split("x"))
                if 400 <= gw <= sw and 300 <= gh <= sh:
                    w, h = gw, gh
            except Exception:
                pass

        x, y = max(0, (sw - w) // 2), max(0, (sh - h) // 3)
        root.geometry("%dx%d+%d+%d" % (w, h, x, y))

    # -- 1-qoida: pastki panellar BIRINCHI pack qilinadi -----------------
    def _build(self):
        self._build_statusbar()     # 1-chi: eng past
        self._build_actionbar()     # 2-chi: status ustida
        self._build_header()        # 3-chi: tepa
        self._build_notebook()      # 4-chi: qolgan joyni egallaydi

    def _build_statusbar(self):
        bar = ttk.Frame(self.root)
        bar.pack(side="bottom", fill="x")
        ttk.Separator(bar, orient="horizontal").pack(fill="x")

        inner = ttk.Frame(bar)
        inner.pack(fill="x", padx=12, pady=(5, 7))

        self.progress = ttk.Progressbar(inner, mode="determinate",
                                        length=int(170 * self.scale))
        self.progress.pack(side="right", padx=(10, 0))

        self.lbl_counts = ttk.Label(inner, text="", style="Muted.TLabel")
        self.lbl_counts.pack(side="right", padx=(10, 0))

        self.lbl_status = ttk.Label(inner, text="Tayyor", style="Muted.TLabel",
                                    anchor="w")
        self.lbl_status.pack(side="left", fill="x", expand=True)

    def _build_actionbar(self):
        """
        Pastki panel - har bo'limdan ko'rinib turadigan IKKITA asosiy amal.
        "Import qilish" va "Qayta hisoblash" endi yo'q: fayl qo'shilishi bilan
        o'zi qabul qilinadi, hisob-kitob hisobotdan oldin o'zi yangilanadi.
        """
        wrap = ttk.Frame(self.root)
        wrap.pack(side="bottom", fill="x")
        ttk.Separator(wrap, orient="horizontal").pack(fill="x")

        bar = FlowBar(wrap, scale=self.scale)
        bar.pack(fill="x", padx=8)
        self.actionbar = bar

        self.btn_add = ttk.Button(bar, text="+  Fayllarni qo'shish", style="BigGhost.TButton",
                                  command=self.pick_files)
        bar.add(self.btn_add, "left")
        self.btn_report = ttk.Button(bar, text="Excel hisobotni yaratish",
                                     style="Big.TButton", command=self.do_report)
        bar.add(self.btn_report, "right")

    def _build_header(self):
        head = ttk.Frame(self.root)
        head.pack(side="top", fill="x", padx=16, pady=(12, 4))
        ttk.Label(head, text="Versiya %s" % C.VERSION, style="Muted.TLabel").pack(side="right", anchor="n")
        left = ttk.Frame(head)
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, text="Buxgalteriya hisoboti", style="H1.TLabel").pack(anchor="w")
        self.lbl_sub = ttk.Label(
            left, style="Muted.TLabel",
            text="Fakturalar va kassa cheklaridan \"KAMERAL TEKSHIRUVLAR\" Excel hisobotini tayyorlaydi")
        self.lbl_sub.pack(anchor="w", pady=(2, 0))

    def _build_notebook(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(side="top", fill="both", expand=True, padx=10, pady=(4, 6))

        self.tab_home = self._tab_home()
        self.tab_files = self.tab_home          # eski nom (testlar va yo'naltirishlar)
        self.tab_match = self._tab_match()
        self.tab_stock = self._tab_stock()
        self.tab_issues = self._tab_issues()
        self.tab_settings = self._tab_settings()

        self.nb.add(self.tab_home, text="  Asosiy  ")
        self.nb.add(self.tab_match, text="  Tovarni tanlash  ")
        self.nb.add(self.tab_stock, text="  Ombor qoldig'i  ")
        self.nb.add(self.tab_issues, text="  Kamchiliklar  ")
        self.nb.add(self.tab_settings, text="  Sozlamalar  ")
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab)

    # ------------------------------------------------------------------
    # Asosiy: 3 qadam
    # ------------------------------------------------------------------
    def _step(self, master, num, title, hint):
        outer, inner = card(master)
        outer.pack(fill="x", pady=(0, 12))
        head = ttk.Frame(inner, style="CardIn.TFrame")
        head.pack(fill="x")
        tk.Label(head, text=str(num), bg=PALETTE["accent"], fg="#ffffff",
                 font=("Segoe UI", self.fonts["h2"][1] + 2, "bold"),
                 width=2).pack(side="left", anchor="n", padx=(0, 12))
        box = ttk.Frame(head, style="CardIn.TFrame")
        box.pack(side="left", fill="x", expand=True)
        ttk.Label(box, text=title, style="Step.TLabel").pack(anchor="w")
        if hint:
            wrapping_label(box, hint, style="CardMuted.TLabel", pad=60).pack(fill="x")
        body = ttk.Frame(inner, style="CardIn.TFrame")
        body.pack(fill="x", pady=(10, 0))
        return body

    def _tab_home(self):
        page = ScrollFrame(self.nb)
        pad = ttk.Frame(page.body, padding=12)
        pad.pack(fill="both", expand=True)

        # ---- 1-qadam -------------------------------------------------
        b1 = self._step(pad, 1, "Fayllarni qo'shing",
                        "Soliq saytidan olingan fakturalar (.xls) va kassa cheklari (checks-info .xlsx). "
                        "Hammasini birdaniga tanlash mumkin. Qo'shilishi bilan o'zi qabul qilinadi. "
                        "Bir fayl ikki marta qo'shilsa - bir marta hisoblanadi.")
        drop = ttk.Frame(b1, style="Drop.TFrame")
        drop.pack(fill="x")
        dz = ttk.Frame(drop, style="DropIn.TFrame")
        dz.pack(fill="x", padx=14, pady=14)
        self.lbl_drop = ttk.Label(
            dz, style="H2.TLabel", anchor="center", justify="center", background="#eef3fb",
            text="Fayllarni shu yerga sudrab tashlang")
        self.lbl_drop.pack(fill="x")
        self.lbl_drop2 = ttk.Label(
            dz, style="CardMuted.TLabel", anchor="center", justify="center", background="#eef3fb",
            text="yoki tugmani bosing:")
        self.lbl_drop2.pack(fill="x", pady=(2, 10))
        dbar = ttk.Frame(dz, style="DropIn.TFrame")
        dbar.pack()
        ttk.Button(dbar, text="Fayllarni tanlash", style="Big.TButton",
                   command=self.pick_files).pack(side="left", padx=6)
        ttk.Button(dbar, text="Papkani tanlash", style="BigGhost.TButton",
                   command=self.pick_folder).pack(side="left", padx=6)

        self._dnd = None
        for w in (self.root, drop):
            b = enable_file_drop(w, self.on_files_dropped)
            self._dnd = self._dnd or b
        if not self._dnd:
            self.lbl_drop.configure(text="Fayllarni tanlang")
            self.lbl_drop2.configure(text="Bir nechtasini Ctrl tugmasini bosib turib tanlash mumkin:")

        self.lbl_import = ttk.Label(b1, text="", style="Result.TLabel", justify="left",
                                    anchor="w")
        self.lbl_import.pack(fill="x", pady=(10, 4))
        self.lbl_queue = ttk.Label(b1, text="", style="CardMuted.TLabel")
        self.lbl_queue.pack(anchor="w")
        self.tbl_files = Table(b1, [
            ("file", "Fayl", 380, True, "w"),
            ("kind", "Turi", 150, False, "w"),
            ("state", "Natija", 330, True, "w"),
        ], scale=self.scale, height=6, on_double=self._toggle_kind)
        self.tbl_files.pack(fill="x")
        self._last_results = []

        # ---- 2-qadam -------------------------------------------------
        b2 = self._step(pad, 2, "Tekshiring: hamma oy bormi?",
                        "Har oyda nechta faktura va chek borligi. Yashil - joyida, sariq - tovar tanlash kerak, "
                        "qizil - fayl qo'shilmagan.")
        yrow = ttk.Frame(b2, style="CardIn.TFrame")
        yrow.pack(fill="x", pady=(0, 8))
        ttk.Label(yrow, text="Yil:", style="Text.TLabel").pack(side="left", padx=(0, 8))
        self.year_bar = ttk.Frame(yrow, style="CardIn.TFrame")
        self.year_bar.pack(side="left")
        self.var_year = tk.IntVar(value=0)
        self.month_grid = ttk.Frame(b2, style="CardIn.TFrame")
        self.month_grid.pack(fill="x")
        self.lbl_months = ttk.Label(b2, text="", style="Result.TLabel", justify="left", anchor="w")
        self.lbl_months.pack(fill="x", pady=(10, 0))
        self.lbl_tin_warn = ttk.Label(b2, text="", style="Err.TLabel", justify="left", anchor="w")
        self.lbl_tin_warn.pack(fill="x")
        self.btn_go_match = ttk.Button(b2, text="Tovarni tanlash  →", style="BigGhost.TButton",
                                       command=lambda: self.nb.select(self.tab_match))

        # ---- 3-qadam -------------------------------------------------
        b3 = self._step(pad, 3, "Excel hisobotni oling",
                        "Tugmani bossangiz hisobot tayyorlanadi, saqlanadi va o'zi ochiladi.")
        self.btn_report_big = ttk.Button(b3, text="Excel hisobotni yaratish", style="Green.TButton",
                                         command=self.do_report)
        self.btn_report_big.pack(anchor="w")
        orow = ttk.Frame(b3, style="CardIn.TFrame")
        orow.pack(fill="x", pady=(10, 0))
        self.var_outdir = tk.StringVar(
            value=DB.get_setting(self.cx, "out_dir",
                                 os.path.join(os.path.expanduser("~"), "Desktop")))
        self.lbl_outdir = ttk.Label(orow, text="", style="CardMuted.TLabel")
        self.lbl_outdir.pack(side="left")
        ttk.Button(orow, text="Boshqa joy...", style="Ghost.TButton",
                   command=self._browse_outdir).pack(side="left", padx=(10, 0))
        self._show_outdir()
        self.lbl_result = ttk.Label(b3, text="", style="Result.TLabel", justify="left", anchor="w")
        self.lbl_result.pack(fill="x", pady=(10, 0))
        rrow = ttk.Frame(b3, style="CardIn.TFrame")
        rrow.pack(fill="x", pady=(6, 0))
        self.btn_open = ttk.Button(rrow, text="Hisobotni ochish", style="Ghost.TButton",
                                   command=self._open_last, state="disabled")
        self.btn_open.pack(side="left")
        self.btn_open_dir = ttk.Button(
            rrow, text="Papkasini ochish", style="Ghost.TButton", state="disabled",
            command=lambda: self._open_path(os.path.dirname(self._last_report or "")))
        self.btn_open_dir.pack(side="left", padx=(8, 0))
        return page

    def _show_outdir(self):
        d = self.var_outdir.get() or os.path.expanduser("~")
        base = os.path.basename(d.rstrip("\\/")).lower()
        name = "Ish stoli" if base in ("desktop", "рабочий стол") else os.path.normpath(d)
        self.lbl_outdir.configure(text="Saqlanadigan joy: %s" % name)

    # ------------------------------------------------------------------
    # Tovarni tanlash
    # ------------------------------------------------------------------
    def _tab_match(self):
        page = ttk.Frame(self.nb, padding=10)

        top, inner = card(page, "Chekdagi tovar fakturadagi qaysi tovar?")
        top.pack(fill="x")
        wrapping_label(
            inner,
            "Kassada tovar nomi boshqacha yozilgani uchun dastur uni o'zi topa olmadi.  "
            "1) Chapdan tovarni bosing.  2) O'ngdan fakturadagi to'g'ri tovarni bosing.  "
            "3) \"Saqlash\".  Keyingi safar o'zi topadi. Fakturasi hali qo'shilmagan bo'lsa - tanlamang."
        ).pack(fill="x")

        pane = ttk.PanedWindow(page, orient="horizontal")
        pane.pack(fill="both", expand=True, pady=(10, 0))

        left = ttk.Frame(pane)
        ttk.Label(left, text="1) Chekda shunday yozilgan:", style="H2.TLabel",
                  background=PALETTE["bg"]).pack(anchor="w", pady=(0, 4))
        self.tbl_unmatched = Table(left, [
            ("name", "Tovar nomi (chekda)", 240, True, "w"),
            ("n", "Necha marta", 115, False, "e"),
            ("qty", "Miqdor", 80, False, "e"),
            ("amount", "Summa", 120, False, "e"),
        ], scale=self.scale, height=14, on_select=self._on_unmatched_select)
        self.tbl_unmatched.pack(fill="both", expand=True)

        right = ttk.Frame(pane)
        rhead = ttk.Frame(right)
        rhead.pack(fill="x", pady=(0, 4))
        ttk.Label(rhead, text="2) Fakturadagi qaysi tovar?", style="H2.TLabel",
                  background=PALETTE["bg"]).pack(side="left")
        # Saqlash tugmasi tepada - kichik ekranda ham doim ko'rinsin
        ttk.Button(rhead, text="3) Saqlash", style="Big.TButton",
                   command=lambda: self.do_link()).pack(side="right")
        self.cand_box = ttk.Frame(right)
        self.cand_box.pack(fill="x")
        self.lbl_nocand = ttk.Label(
            self.cand_box, style="Muted.TLabel", justify="left",
            text="Chapdan tovarni tanlang. Mos tovar topilmasa - pastdagi qidiruvdan toping.")
        self.lbl_nocand.pack(anchor="w", pady=4)
        self.tbl_cand = Table(self.cand_box, [
            ("name", "Mos keladigan tovarlar", 230, True, "w"),
            ("score", "O'xshashlik", 115, False, "e"),
            ("why", "Nega", 150, False, "w"),
            ("left", "Qoldiq", 75, False, "e"),
        ], scale=self.scale, height=4, on_double=lambda e: self.do_link("cand"))

        srow = ttk.Frame(right)
        srow.pack(fill="x", pady=(8, 4))
        ttk.Label(srow, text="Boshqasini qidirish:").pack(side="left")
        self.var_search = tk.StringVar()
        ent = ttk.Entry(srow, textvariable=self.var_search)
        ent.pack(side="left", fill="x", expand=True, padx=6)
        ent.bind("<KeyRelease>", lambda e: self._search_products())

        self.tbl_prod = Table(right, [
            ("name", "Hamma fakturadagi tovarlar", 230, True, "w"),
            ("mxik", "MXIK", 130, False, "w"),
            ("left", "Qoldiq", 70, False, "e"),
        ], scale=self.scale, height=5, on_double=lambda e: self.do_link("prod"))
        self.tbl_prod.pack(fill="both", expand=True)

        pane.add(left, weight=3)
        pane.add(right, weight=4)
        return page

    # ------------------------------------------------------------------
    # Ombor qoldig'i
    # ------------------------------------------------------------------
    def _tab_stock(self):
        page = ttk.Frame(self.nb, padding=10)
        row = ttk.Frame(page)
        row.pack(fill="x", pady=(0, 8))
        ttk.Label(row, text="Omborda nima qoldi", style="H1.TLabel").pack(side="left")
        self.var_stock_zero = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="Qoldig'i yo'qlarini ham ko'rsatish",
                        variable=self.var_stock_zero,
                        command=self.refresh_stock).pack(side="right")

        self.tbl_stock = Table(page, [
            ("name", "Tovar", 280, True, "w"),
            ("unit", "Birlik", 90, False, "w"),
            ("inq", "Kelgan", 85, False, "e"),
            ("outq", "Sotilgan", 85, False, "e"),
            ("left", "Qoldiq", 85, False, "e"),
            ("cost", "Tannarx (o'rtacha)", 150, False, "e"),
            ("value", "Qoldiq summasi", 150, False, "e"),
        ], scale=self.scale, height=16)
        self.tbl_stock.pack(fill="both", expand=True)
        ttk.Label(page, style="Muted.TLabel",
                  text="Qizil satr: fakturadagidan ko'p sotilgan - shu tovarning fakturasi yetishmaydi."
                  ).pack(anchor="w", pady=(4, 0))
        return page

    # ------------------------------------------------------------------
    # Kamchiliklar
    # ------------------------------------------------------------------
    def _tab_issues(self):
        page = ttk.Frame(self.nb, padding=10)
        row = ttk.Frame(page)
        row.pack(fill="x", pady=(0, 8))
        ttk.Label(row, text="Nimani tekshirish kerak", style="H1.TLabel").pack(side="left")
        self.var_sev = tk.StringVar(value="hammasi")
        cb = ttk.Combobox(row, textvariable=self.var_sev, state="readonly", width=18,
                          values=["hammasi", C.SEVERITY_ERROR, C.SEVERITY_WARN,
                                  C.SEVERITY_INFO])
        cb.pack(side="right")
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_issues())
        ttk.Label(row, text="Ko'rsatish:", style="Muted.TLabel").pack(side="right", padx=6)

        self.tbl_issues = Table(page, [
            ("sev", "Darajasi", 130, False, "w"),
            ("year", "Yil", 60, False, "e"),
            ("code", "Nima", 240, False, "w"),
            ("msg", "Batafsil", 400, True, "w"),
        ], scale=self.scale, height=16)
        self.tbl_issues.pack(fill="both", expand=True)
        ttk.Label(page, style="Muted.TLabel",
                  text="Qizil - albatta ko'rib chiqing. Sariq - ma'lumot uchun. Hisobot baribir yaratiladi."
                  ).pack(anchor="w", pady=(4, 0))
        return page

    # ------------------------------------------------------------------
    # Sozlamalar (kamdan-kam kerak bo'ladigan narsalar)
    # ------------------------------------------------------------------
    def _tab_settings(self):
        page = ScrollFrame(self.nb)
        pad = ttk.Frame(page.body, padding=10)
        pad.pack(fill="both", expand=True)

        c0, i0 = card(pad, "Hisobot sarlavhasi")
        c0.pack(fill="x")
        row = ttk.Frame(i0, style="CardIn.TFrame")
        row.pack(fill="x", pady=3)
        ttk.Label(row, text="Tashkilot nomi", style="Card.TLabel", width=22).pack(side="left")
        self.var_owner = tk.StringVar(value=DB.get_setting(self.cx, "owner_name", ""))
        ttk.Entry(row, textvariable=self.var_owner).pack(side="left", fill="x", expand=True)

        c1, i1 = card(pad, "Ustama (yil bo'yicha)")
        c1.pack(fill="x", pady=(10, 0))
        wrapping_label(
            i1, "Sotish narxi = tannarx x (1 + ustama) x (1 + QQS). Odatda o'zgartirish shart emas."
        ).pack(fill="x", pady=(0, 8))
        self.markup_vars = {}
        grid = ttk.Frame(i1, style="CardIn.TFrame")
        grid.pack(fill="x")
        for i, y in enumerate(range(2023, datetime.date.today().year + 2)):
            v = tk.StringVar(value=("%g" % (DB.get_markup(self.cx, y) * 100)))
            self.markup_vars[y] = v
            cell = ttk.Frame(grid, style="CardIn.TFrame")
            cell.grid(row=i // 3, column=i % 3, sticky="w", padx=(0, 22), pady=4)
            ttk.Label(cell, text="%d:" % y, style="Card.TLabel", width=7).pack(side="left")
            ttk.Entry(cell, textvariable=v, width=8).pack(side="left")
            ttk.Label(cell, text="%", style="Card.TLabel").pack(side="left", padx=(4, 0))
        ttk.Button(i1, text="Saqlash", style="Ghost.TButton",
                   command=self.save_markups).pack(anchor="w", pady=(10, 0))

        top, inner = card(pad, "Doimiy papkalar (shart emas)")
        top.pack(fill="x", pady=(10, 0))
        wrapping_label(
            inner,
            "Fayllar doim bir papkada tursa, shu yerda belgilab qo'ying - dastur har ochilganda "
            "o'sha papkadagi yangi fayllarni o'zi qo'shadi."
        ).pack(fill="x", pady=(0, 10))
        self.var_fak = tk.StringVar(value=DB.get_setting(self.cx, "folder_kirim", ""))
        self.var_chq = tk.StringVar(value=DB.get_setting(self.cx, "folder_chiqim", ""))
        for label, var, kind in (("Fakturalar papkasi", self.var_fak, "kirim"),
                                 ("Kassa cheklari papkasi", self.var_chq, "chiqim")):
            row = ttk.Frame(inner, style="CardIn.TFrame")
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=label, style="Card.TLabel", width=22).pack(side="left")
            ttk.Button(row, text="Tozalash", style="Ghost.TButton",
                       command=lambda v=var, k=kind: self._clear_folder(v, k)
                       ).pack(side="right", padx=(6, 0))
            ttk.Button(row, text="Tanlash", style="Ghost.TButton",
                       command=lambda v=var, k=kind: self._browse_folder(v, k)
                       ).pack(side="right", padx=(8, 0))
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

        c5, i5 = card(pad, "Hisob-kitobni qaytadan bajarish")
        c5.pack(fill="x", pady=(10, 0))
        wrapping_label(
            i5, "Odatda kerak emas - hisobot oldidan o'zi bajariladi. Ustamani o'zgartirgandan keyin bosing."
        ).pack(fill="x", pady=(0, 8))
        self.btn_recalc = ttk.Button(i5, text="Qayta hisoblash", style="Ghost.TButton",
                                     command=self.do_recalc)
        self.btn_recalc.pack(anchor="w")

        c2, i2 = card(pad, "Ma'lumotlar bazasi")
        c2.pack(fill="x", pady=(10, 0))
        self.lbl_db = ttk.Label(i2, style="CardMuted.TLabel", justify="left", anchor="w")
        self.lbl_db.pack(fill="x")
        brow = FlowBar(i2, scale=self.scale)
        brow.pack(fill="x", pady=(8, 0))
        brow.add(ttk.Button(brow, text="Zaxira nusxa olish", style="Ghost.TButton",
                            command=self.do_backup), "left")
        brow.add(ttk.Button(brow, text="Papkani ochish", style="Ghost.TButton",
                            command=lambda: self._open_path(C.app_data_dir())), "left")
        brow.add(ttk.Button(brow, text="Hammasini o'chirish", style="Ghost.TButton",
                            command=self.do_reset), "left")

        c4, i4 = card(pad, "Ish stolidagi belgi")
        c4.pack(fill="x", pady=(10, 0))
        wrapping_label(
            i4, "Ish stolidagi belgi o'chib ketgan bo'lsa - shu tugmani bosing."
        ).pack(fill="x", pady=(0, 8))
        ttk.Button(i4, text="Ish stoliga belgi qo'yish", style="Ghost.TButton",
                   command=self.do_shortcut).pack(anchor="w")

        c3, i3 = card(pad, "Dastur haqida (dasturchi uchun)")
        c3.pack(fill="x", pady=(10, 0))
        origins = self.info.get("origins") or {}
        commit = self.info.get("commit")
        txt = ["Versiya:       %s" % C.VERSION,
               "Launcher:      %s" % self.info.get("launcher_version", "-"),
               "Yangilanish:   %s" % self.info.get("base_url", "-"),
               "Commit:        %s" % (commit[:12] if commit else "-"),
               "Baza:          %s" % C.db_path(),
               ""]
        for k, v in origins.items():
            txt.append("  %-14s %s" % (k, v))
        ttk.Label(i3, text="\n".join(txt), style="CardMuted.TLabel",
                  justify="left", anchor="w", font=self.fonts["mono"]).pack(fill="x")
        return page

    # ==================================================================
    # Yangilash
    # ==================================================================
    def refresh_all(self):
        """
        Yengil qismlar darrov, og'irlari esa bo'lim ochilganda.

        Ilgari har amaldan keyin hamma jadval qaytadan hisoblanardi - shu
        jumladan har yil uchun to'liq FIFO kesimi (year_rows). Bu interfeys
        oqimida bajarilgani uchun oyna bir necha soniya qotib turardi,
        holbuki foydalanuvchi ko'pincha bitta bo'limga qaraydi.
        """
        self.refresh_counts()
        self.refresh_queue()
        self._refresh_db_label()
        self._stale = {0, 1, 2, 3}
        self._refresh_tab(self._current_tab())

    def _current_tab(self):
        try:
            return self.nb.index(self.nb.select())
        except tk.TclError:
            return -1

    def _refresh_tab(self, tab):
        """Bitta bo'limni yangilaydi (kerak bo'lsa)."""
        if tab not in self._stale:
            return
        self._stale.discard(tab)
        if tab == 0:
            self.refresh_home()
        elif tab == 1:
            self.refresh_unmatched()
            self._search_products()
        elif tab == 2:
            self.refresh_stock()
        elif tab == 3:
            self.refresh_issues()

    def refresh_counts(self):
        st = DB.stats(self.cx)
        self.lbl_counts.configure(
            text="Faktura: %d  ·  Chek: %d  ·  Tovar tanlash kerak: %d"
                 % (st["docs_in"], st["docs_out"], st["unmatched"]))

    def _file_info(self, path, stt):
        """
        Faylning sha256 i va formati; (yo'l, hajm, o'zgargan vaqt) bo'yicha
        keshlanadi.

        Ilgari navbat har yangilanganda (fayl qo'shilganda, tur almashganda,
        har amaldan keyin) HAR BIR fayl boshidan oxirigacha qayta o'qilardi.
        O'nlab fayl tashlansa oyna shu paytda qotib turardi.
        """
        key = (os.path.normcase(os.path.abspath(path)), stt.st_size,
               int(stt.st_mtime))
        got = self._sha_cache.get(key)
        if got is None:
            got = (DB.file_sha256(path), P.sniff(path))
            self._sha_cache[key] = got
        return got

    KIND_WORD = {"kirim": "Faktura", "chiqim": "Kassa cheki", "legacy": "Eski hisobot",
                 "nomalum": "Noma'lum"}

    def refresh_queue(self):
        """
        Jadvalda: navbatdagi fayllar (qabul qilinmoqda) yoki oxirgi qo'shilganlar natijasi.
        """
        rows = []
        for p, kind in self.queued_files:
            try:
                stt = os.stat(p)
                sha, _fmt = self._file_info(p, stt)
                ex = DB.find_source_file(self.cx, sha)
                if ex is None or F.needs_reparse(self.cx, ex):
                    state = "Navbatda - qabul qilinadi"
                else:
                    state = "Oldin qo'shilgan (%s)" % (ex["imported_at"] or "")[:10]
            except OSError:
                state = "Fayl ochilmadi"
            rows.append((os.path.basename(p), self.KIND_WORD.get(kind, kind), state))
        if not rows:
            rows = list(self._last_results)

        def tag(r):
            st = r[2]
            if st.startswith("Oldin"):
                return ("warn",)
            if st.startswith("Fayl ochilmadi") or st.startswith("O'qilmadi"):
                return ("err",)
            if st.startswith("Qabul qilindi"):
                return ("ok",)
            return ()

        self.tbl_files.fill(rows, tag)
        if self.queued_files:
            self.lbl_queue.configure(text="%d ta fayl navbatda" % len(self.queued_files))
        elif self._last_results:
            self.lbl_queue.configure(text="Oxirgi qo'shilgan fayllar:")
        else:
            self.lbl_queue.configure(text="")

    def refresh_unmatched(self):
        self.engine.reload()
        rows = self.engine.unmatched_groups()
        self._unmatched = [dict(r) for r in rows]
        self.tbl_unmatched.fill(
            [((r["raw_name"] or "")[:90], r["n"], fmt_qty(r["qty"]),
              fmt_money(r["amount"])) for r in self._unmatched],
            lambda r: ("err",) if False else ())

    def refresh_stock(self):
        show_zero = self.var_stock_zero.get()
        rows = self.cx.execute("""
            SELECT p.id, p.canon_name, p.unit,
                   COALESCE(SUM(CAST(s.qty_in AS REAL)), 0) inq,
                   COALESCE(SUM(CAST(s.qty_left AS REAL)), 0) leftq,
                   CASE WHEN SUM(CAST(s.qty_in AS REAL)) > 0
                        THEN SUM(CAST(s.qty_in AS REAL) * CAST(s.unit_cost AS REAL))
                             / SUM(CAST(s.qty_in AS REAL))
                        ELSE 0 END avg_cost
            FROM product p LEFT JOIN stock_lot s ON s.product_id = p.id
            GROUP BY p.id ORDER BY p.canon_name""").fetchall()
        out = []
        for r in rows:
            left = r["leftq"] or 0
            if not show_zero and abs(left) < 1e-9:
                continue
            out.append((r["canon_name"][:90], r["unit"] or "",
                        fmt_qty(r["inq"]), fmt_qty((r["inq"] or 0) - left),
                        fmt_qty(left), fmt_money(r["avg_cost"]),
                        fmt_money(left * (r["avg_cost"] or 0))))

        def tag(row):
            try:
                return ("err",) if float(str(row[4]).replace(" ", "")) < 0 else ()
            except ValueError:
                return ()

        self.tbl_stock.fill(out, tag)

    def refresh_issues(self):
        sev = self.var_sev.get()
        sql = "SELECT * FROM issue WHERE resolved=0"
        args = []
        if sev != "hammasi":
            sql += " AND severity=?"
            args.append(sev)
        sql += (" ORDER BY CASE severity WHEN 'xato' THEN 0 "
                "WHEN 'ogohlantirish' THEN 1 ELSE 2 END, year DESC, id DESC LIMIT 3000")
        rows = self.cx.execute(sql, args).fetchall()
        self.tbl_issues.fill(
            [(r["severity"], r["year"] or "",
              C.ISSUE_TITLES.get(r["code"], r["code"]), r["message"])
             for r in rows],
            lambda r: ("err",) if r[0] == C.SEVERITY_ERROR else
                      (("warn",) if r[0] == C.SEVERITY_WARN else ()))

    MONTHS = ["Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun", "Iyul", "Avgust",
              "Sentabr", "Oktabr", "Noyabr", "Dekabr"]

    def refresh_home(self):
        """2-qadam: yil tanlovi, 12 oy katakchasi, oddiy so'z bilan xulosa."""
        years = DB.available_years(self.cx)
        for w in self.year_bar.winfo_children():
            w.destroy()
        for w in self.month_grid.winfo_children():
            w.destroy()
        self.btn_go_match.pack_forget()
        if not years:
            ttk.Label(self.year_bar, text="hali fayl qo'shilmagan",
                      style="CardMuted.TLabel").pack(side="left")
            self.lbl_months.configure(text="Avval 1-qadamda fayllarni qo'shing.")
            self.lbl_tin_warn.configure(text="")
            return
        if self.var_year.get() not in years:
            self.var_year.set(years[-1])
        for y in years:
            ttk.Radiobutton(self.year_bar, text=str(y), value=y, variable=self.var_year,
                            command=self.refresh_home).pack(side="left", padx=(0, 14))
        y = self.var_year.get()

        cnt = {}
        for r in self.cx.execute(
                "SELECT kind, substr(doc_date,6,2) m, COUNT(*) n FROM document "
                "WHERE doc_year=? AND doc_date IS NOT NULL GROUP BY kind, m", (y,)):
            cnt[(r["kind"], r["m"])] = r["n"]
        unm, kas = {}, {}
        for r in self.cx.execute(
                "SELECT substr(d.doc_date,6,2) m, SUM(l.product_id IS NULL) u, "
                "SUM(CAST(l.amount_gross AS REAL)) g FROM doc_line l "
                "JOIN document d ON d.id=l.document_id "
                "WHERE l.kind='chiqim' AND d.doc_year=? GROUP BY m", (y,)):
            unm[r["m"]] = r["u"] or 0
            kas[r["m"]] = r["g"] or 0

        colors = {"ok": (PALETTE["ok_bg"], PALETTE["ok"]),
                  "warn": (PALETTE["warn_bg"], PALETTE["warn"]),
                  "bad": (PALETTE["err_bg"], PALETTE["err"]),
                  "empty": (PALETTE["grey_bg"], PALETTE["muted"])}
        fb = self.fonts["base"][1]
        have = []
        for i, name in enumerate(self.MONTHS):
            m = "%02d" % (i + 1)
            nf, nc, nu = cnt.get(("kirim", m), 0), cnt.get(("chiqim", m), 0), unm.get(m, 0)
            if not nf and not nc:
                st, txt = "empty", "Ma'lumot yo'q"
            elif not nc:
                st, txt = "bad", "Chek qo'shilmagan"
            elif not nf:
                st, txt = "bad", "Faktura qo'shilmagan"
            elif nu:
                st, txt = "warn", "%d ta sotuvga tovar tanlash kerak" % nu
            else:
                st, txt = "ok", "Joyida"
            if st != "empty":
                have.append(name)
            bg, fg = colors[st]
            cell = tk.Frame(self.month_grid, bg=bg, highlightbackground=fg, highlightthickness=2)
            cell.grid(row=i // 4, column=i % 4, sticky="nsew", padx=4, pady=4)
            tk.Label(cell, text=name, bg=bg, fg=fg, font=("Segoe UI", fb + 3, "bold"),
                     anchor="w").pack(fill="x", padx=10, pady=(8, 0))
            tk.Label(cell, text="Faktura: %d ta\nChek: %d ta" % (nf, nc), bg=bg,
                     fg=PALETTE["ink"], font=("Segoe UI", fb), justify="left",
                     anchor="w").pack(fill="x", padx=10)
            if kas.get(m):
                tk.Label(cell, text="Tushum: %s so'm" % fmt_money(kas[m]).rsplit(".", 1)[0],
                         bg=bg, fg=PALETTE["muted"], font=("Segoe UI", fb - 1),
                         anchor="w").pack(fill="x", padx=10)
            tk.Label(cell, text=txt, bg=bg, fg=fg, font=("Segoe UI", fb, "bold"), anchor="w",
                     wraplength=int(230 * self.scale), justify="left").pack(
                fill="x", padx=10, pady=(2, 8))
        for c in range(4):
            self.month_grid.columnconfigure(c, weight=1, uniform="m")

        total_u = sum(unm.values())
        lines = []
        if have:
            lines.append("Ma'lumot bor oylar: %s%s." % (
                have[0], (" - %s" % have[-1]) if len(have) > 1 else ""))
        if total_u:
            ng = self.cx.execute(
                "SELECT COUNT(DISTINCT l.norm_name) FROM doc_line l "
                "JOIN document d ON d.id=l.document_id "
                "WHERE l.kind='chiqim' AND l.product_id IS NULL AND d.doc_year=?",
                (y,)).fetchone()[0]
            lines.append(
                "%d xil tovarning fakturasi topilmadi (%d ta sotuv). Hisobot baribir chiqadi - "
                "bu satrlar Excelda qizil bo'ladi. Tuzatish uchun fakturasini qo'shing yoki "
                "to'g'ri tovarni tanlang." % (ng, total_u))
            self.btn_go_match.pack(anchor="w", pady=(8, 0))
        elif have:
            lines.append("Hamma sotuv fakturadagi tovarga bog'langan.")
        self.lbl_months.configure(text="\n".join(lines), wraplength=int(900 * self.scale))

        tins = [r[0] for r in self.cx.execute(
            "SELECT DISTINCT partner_tin FROM document WHERE kind='chiqim' AND doc_year=? "
            "AND partner_tin IS NOT NULL AND partner_tin<>''", (y,))]
        if len(tins) > 1:
            self.lbl_tin_warn.configure(
                wraplength=int(900 * self.scale),
                text="Diqqat: bazada %d ta tashkilotning (STIR: %s) cheklari aralash. Hisobot "
                     "bitta bo'lib chiqadi. Har tashkilot uchun alohida hisobot kerak bo'lsa - "
                     "Sozlamalar > \"Hammasini o'chirish\" ni bosib, faqat bitta tashkilot "
                     "fayllarini qo'shing." % (len(tins), ", ".join(tins)))
        else:
            self.lbl_tin_warn.configure(text="")

    def _refresh_db_label(self):
        p = C.db_path()
        size = os.path.getsize(p) / 1024.0 if os.path.exists(p) else 0
        self.lbl_db.configure(
            text="Joylashuv: %s\nHajmi: %.0f KB\nZaxira nusxalar: %s"
                 % (p, size, os.path.join(C.app_data_dir(), "backup")))

    # ==================================================================
    # Fayl tanlash
    # ==================================================================
    def _browse_folder(self, var, kind):
        d = filedialog.askdirectory(title="Papkani tanlang", mustexist=True,
                                    initialdir=var.get() or os.path.expanduser("~"))
        if d:
            var.set(d)
            DB.set_setting(self.cx, "folder_%s" % kind, d)
            self.scan_folders()

    def _browse_outdir(self):
        d = filedialog.askdirectory(title="Saqlash papkasi", mustexist=True,
                                    initialdir=self.var_outdir.get())
        if d:
            self.var_outdir.set(d)
            DB.set_setting(self.cx, "out_dir", d)
            self._show_outdir()

    def pick_folder(self):
        d = filedialog.askdirectory(title="Fayllar papkasini tanlang", mustexist=True)
        if d:
            self._accept([d], source="papka")

    def pick_files(self):
        """Bitta fayl ham, o'nlab fayl ham - Ctrl/Shift bilan ko'p tanlash."""
        fs = filedialog.askopenfilenames(
            title="Excel fayllarini tanlang (bir nechtasini Ctrl bilan)",
            filetypes=[("Excel / faktura", "*.xls *.xlsx *.xlsm *.htm *.html"),
                       ("Barcha fayllar", "*.*")])
        if fs:
            self._accept(list(fs), source="tanlov")

    def on_files_dropped(self, paths):
        """Explorer'dan sudrab tashlanganda."""
        self._accept(paths, source="tashlandi")

    def _accept(self, paths, source=""):
        """
        Kiruvchi yo'llarni navbatga qo'shadi.

        Fayl, bir necha fayl, papka yoki aralash - farqi yo'q. Hammasi
        bitta navbatga tushadi va bitta hisobotga qo'shiladi.
        """
        files = expand_inputs(paths)
        if not files:
            messagebox.showinfo(
                "Mos fayl yo'q",
                "Excel yoki faktura fayli topilmadi.\n\n"
                "Qo'llab-quvvatlanadigan kengaytmalar: %s"
                % ", ".join(P.SUPPORTED_EXT))
            return
        before = len(self.queued_files)
        for f in files:
            self._queue(f, P.guess_kind(f))
        added = len(self.queued_files) - before
        self.nb.select(self.tab_home)
        self.refresh_queue()
        self.set_status("%d ta fayl qo'shildi - qabul qilinmoqda..." % added)
        self._schedule_import()

    def _schedule_import(self):
        """
        Qo'shilgan fayllar darhol qabul qilinadi - buxgalter alohida "Import"
        tugmasini bosishni eslab o'tirmasin. Bir necha marta tez qo'shilsa ham
        bitta ish bo'lib ketadi.
        """
        if getattr(self, "_import_pending", False):
            return
        self._import_pending = True

        def go():
            self._import_pending = False
            if not self.queued_files:
                return
            if self.worker.busy:
                self.root.after(500, self._schedule_import)
                return
            self.do_import()
        self.root.after(150, go)

    def _queue(self, path, kind):
        key = os.path.normcase(os.path.abspath(path))
        if not any(os.path.normcase(os.path.abspath(p)) == key
                   for p, _k in self.queued_files):
            self.queued_files.append((path, kind))

    def _toggle_kind(self, _e=None):
        """Navbat satriga ikki marta bosilsa kirim <-> chiqim almashadi."""
        i = self.tbl_files.selected_index()
        if i is None or i >= len(self.queued_files):
            return
        path, kind = self.queued_files[i]
        self.queued_files[i] = (path, "chiqim" if kind == "kirim" else "kirim")
        self.refresh_queue()
        self.tbl_files.tree.selection_set(str(i))
        self.set_status("%s -> %s" % (os.path.basename(path),
                                      self.queued_files[i][1]))

    def remove_queued(self):
        i = self.tbl_files.selected_index()
        if i is None or i >= len(self.queued_files):
            messagebox.showinfo("Tanlanmagan", "Avval navbatdan satrni tanlang.")
            return
        path, _k = self.queued_files.pop(i)
        self.refresh_queue()
        self.set_status("Navbatdan olindi: %s" % os.path.basename(path))

    def _clear_folder(self, var, kind):
        var.set("")
        DB.set_setting(self.cx, "folder_%s" % kind, "")
        self.set_status("Doimiy papka o'chirildi")

    def clear_queue(self):
        self.queued_files = []
        self.refresh_queue()

    def scan_folders(self):
        """
        Doimiy papkalarni skanerlaydi.

        Navbatni TOZALAMAYDI - qo'lda tashlangan fayllar joyida qoladi va
        papkadagilar ustiga qo'shiladi.
        """
        before = len(self.queued_files)
        found = 0
        for var, kind in ((self.var_fak, "kirim"), (self.var_chq, "chiqim")):
            d = var.get().strip()
            if d and os.path.isdir(d):
                for f in P.scan_folder(d):
                    found += 1
                    self._queue(f, kind)   # papka turi qo'lda o'rnatilgan turdan ustun
        self.refresh_queue()
        added = len(self.queued_files) - before
        if found:
            self.set_status("Doimiy papkalarda %d ta fayl" % found)
            if added:
                self._schedule_import()
        else:
            self.set_status("Doimiy papkalarda fayl topilmadi")

    # ==================================================================
    # Amallar
    # ==================================================================
    def do_import(self):
        files = list(self.queued_files)
        if not files:
            messagebox.showinfo("Bo'sh", "Import uchun fayl tanlanmagan.")
            return

        def job(cx, progress):
            eng = M.MatchEngine(cx)
            docs = lines = skipped = dup_files = 0
            warns = []
            results = []       # (fayl, turi, natija) - jadval uchun
            word = self.KIND_WORD
            for i, (path, kind) in enumerate(files):
                name = os.path.basename(path)
                progress(i, len(files), "Qabul qilinmoqda: %d / %d" % (i + 1, len(files)))
                try:
                    sid, isnew = DB.add_source_file(cx, path, kind)
                except OSError as e:
                    warns.append("%s: %s" % (name, e))
                    results.append((name, word.get(kind, kind), "Fayl ochilmadi"))
                    continue
                if not isnew:
                    sf = cx.execute("SELECT * FROM source_file WHERE id=?",
                                    (sid,)).fetchone()
                    if F.needs_reparse(cx, sf):
                        # eski versiya to'liq o'qimagan chek fayli
                        a, n, w = F.reparse_file(cx, eng, sf, path)
                        warns.extend(w)
                        docs += a
                        lines += n
                        results.append((name, word.get(kind, kind),
                                        "Qabul qilindi: qayta o'qildi, %d ta yangi hujjat" % a))
                    else:
                        dup_files += 1
                        results.append((name, word.get(kind, kind),
                                        "Oldin qo'shilgan - qayta hisoblanmaydi"))
                    continue
                r = P.parse_any(path, kind)
                warns.extend(r["warnings"])
                a, s, n = F.import_parsed(cx, eng, r, sid)
                DB.update_source_counts(cx, sid, a, n)
                docs += a
                skipped += s
                lines += n
                if a:
                    what = "faktura" if kind == "kirim" else "chek" if kind == "chiqim" else "hujjat"
                    results.append((name, word.get(kind, kind),
                                    "Qabul qilindi: %d ta %s, %d ta satr" % (a, what, n)))
                elif s:
                    results.append((name, word.get(kind, kind),
                                    "Oldin qo'shilgan - ichidagi hujjatlar bazada bor"))
                else:
                    why = (r["warnings"][0].split(": ", 1)[-1] if r["warnings"] else
                           "ichida hujjat topilmadi")
                    results.append((name, word.get(kind, kind), "O'qilmadi: %s" % why))
            old = F.reparse_old_checks(cx, eng, progress)
            docs += old["docs"]
            lines += old["lines"]
            progress(len(files), len(files), "moslashtirish")
            eng.reload()
            auto, left = M.auto_match_all(cx, eng, progress=progress)
            progress(0, 0, "ombor hisoblanmoqda")
            st = F.rebuild_stock(cx, progress)
            return {"files": len(files), "dup_files": dup_files, "docs": docs,
                    "skipped": skipped, "lines": lines, "auto": auto,
                    "left": left, "stock": st, "warns": warns, "results": results}

        self._run("Import", job)

    def do_recalc(self):
        def job(cx, progress):
            eng = M.MatchEngine(cx)
            F.reparse_old_checks(cx, eng, progress)
            eng.reload()
            progress(0, 0, "moslashtirish")
            auto, left = M.auto_match_all(cx, eng, progress=progress)
            progress(0, 0, "ombor (FIFO)")
            st = F.rebuild_stock(cx, progress)
            for y in DB.available_years(cx):
                F.validate_year(cx, y)
            return {"auto": auto, "left": left, "stock": st}

        self._run("Qayta hisoblash", job)

    def do_report(self):
        years = DB.available_years(self.cx)
        if not years:
            messagebox.showwarning(
                "Ma'lumot yo'q",
                "Hisobot uchun avval fayllarni qo'shing (1-qadam).")
            self.nb.select(self.tab_home)
            return

        outdir = self.var_outdir.get().strip() or os.path.expanduser("~")
        if not os.path.isdir(outdir):
            outdir = os.path.expanduser("~")
        owner = self.var_owner.get().strip() or "Ташкилот"
        DB.set_setting(self.cx, "owner_name", owner)
        DB.set_setting(self.cx, "out_dir", outdir)

        # Saqlash oynasi so'ralmaydi: buxgalter fayl nomi va papka bilan
        # ovora bo'lmasin. Bir xil nomli fayl bo'lsa (yoki Excelda ochiq
        # bo'lsa) - yoniga (2), (3) qo'shiladi.
        base = R.default_filename(years)
        stem, ext = os.path.splitext(base)
        path = os.path.join(outdir, base)
        n = 2
        while os.path.exists(path):
            path = os.path.join(outdir, "%s (%d)%s" % (stem, n, ext))
            n += 1

        self.lbl_result.configure(text="Tayyorlanmoqda... biroz kuting.")

        def job(cx, progress):
            DB.set_setting(cx, "owner_name", owner)
            p, st = R.generate(cx, path, years=years, owner_name=owner,
                               progress=lambda a, b, m: progress(a, b, m))
            return {"path": p, "stats": st}

        self._run("Hisobot", job)

    def do_link(self, which=None):
        ui = self.tbl_unmatched.selected_index()
        if ui is None or ui >= len(self._unmatched):
            messagebox.showinfo("Tanlanmagan",
                                "Avval chapdagi ro'yxatdan nomni tanlang.")
            return

        pid = None
        if which in (None, "cand"):
            ci = self.tbl_cand.selected_index()
            if ci is not None and ci < len(self._candidates):
                pid = self._candidates[ci][0]
        if pid is None and which in (None, "prod"):
            pi = self.tbl_prod.selected_index()
            if pi is not None and pi < len(self._products):
                pid = self._products[pi]["id"]
        if pid is None:
            messagebox.showinfo(
                "Tanlanmagan",
                "O'ngdagi ro'yxatdan mos kirim mahsulotini tanlang "
                "(ustiga ikki marta bosish ham bog'laydi).")
            return

        grp = self._unmatched[ui]
        self.engine.confirm(grp["line_id"], pid, grp["raw_name"])
        name = self.engine.products.get(pid, {}).get("canon_name", "?")
        self._stock_dirty = True
        self._stale.update({0, 2, 3})
        self.set_status("Bog'landi: %s  ->  %s" % ((grp["raw_name"] or "")[:40],
                                                   name[:40]))
        self.refresh_unmatched()
        self.refresh_counts()
        if not self._unmatched:
            messagebox.showinfo(
                "Tugadi",
                "Hamma sotuv fakturadagi tovarga bog'landi.\n\n"
                "Endi \"Asosiy\" bo'limida Excel hisobotni yaratishingiz mumkin.")

    def save_markups(self):
        ok = 0
        for y, v in self.markup_vars.items():
            try:
                pct = float(str(v.get()).replace(",", ".").strip())
            except ValueError:
                messagebox.showerror("Xato", "%d yil uchun noto'g'ri qiymat: %s"
                                     % (y, v.get()))
                return
            DB.set_markup_for_year(self.cx, y, str(pct / 100.0))
            ok += 1
        messagebox.showinfo(
            "Saqlandi",
            "%d yil uchun ustama saqlandi.\n\nYangi narxlar kuchga kirishi "
            "uchun \"Qayta hisoblash\" tugmasini bosing." % ok)

    def do_shortcut(self):
        exe = app_exe_path()
        if not exe:
            messagebox.showinfo(
                "Ishlab chiqish rejimi",
                "Yorliq faqat .exe ishga tushirilganda yaratiladi.\n\n"
                "Hozir dastur to'g'ridan-to'g'ri Python orqali ishlayapti.")
            return
        made = install_shortcuts()
        if made:
            messagebox.showinfo("Tayyor",
                                "Yorliq yaratildi:\n\n%s" % "\n".join(made))
        else:
            messagebox.showerror(
                "Bajarilmadi",
                "Yorliq yaratilmadi. Antivirus yoki tashkilot siyosati "
                "to'sgan bo'lishi mumkin.\n\n"
                "Qo'lda: %s faylini o'ng tugma bilan bosib "
                "\"Ish stoliga yuborish\" ni tanlang." % exe)

    def do_backup(self):
        p = DB.backup(C.db_path())
        if p:
            messagebox.showinfo("Tayyor", "Zaxira nusxa:\n\n%s" % p)
            self._refresh_db_label()
        else:
            messagebox.showerror("Xato", "Zaxira nusxa olinmadi.")

    def do_reset(self):
        if not messagebox.askyesno(
                "Tasdiqlang",
                "Barcha import qilingan hujjat va hisob-kitob o'chiriladi.\n\n"
                "Mahsulot kartochkalari va o'rganilgan mosliklar SAQLANADI.\n\n"
                "Davom etamizmi?"):
            return
        DB.backup(C.db_path(), suffix="reset")
        DB.reset_all(self.cx)
        self.queued_files = []
        self.refresh_all()
        self.set_status("Ma'lumot o'chirildi (zaxira nusxa olindi)")

    # ==================================================================
    # Yordamchi
    # ==================================================================
    def _on_tab(self, _e=None):
        try:
            tab = self.nb.index(self.nb.select())
        except tk.TclError:
            return
        if self._recalc_if_dirty():
            return
        if tab in self._stale:
            self._refresh_tab(tab)
        elif tab == 1:
            self.refresh_unmatched()
        elif tab == 0:
            self.refresh_home()

    def _recalc_if_dirty(self):
        """
        Qo'lda bog'langandan keyin omborni o'zi qayta hisoblaydi.

        Buxgalter "Qayta hisoblash" tugmasini bosishni eslab o'tirmasin:
        "Ombor" yoki "Hisobot" bo'limi ochilganda raqamlar o'zi yangilanadi.
        """
        if not self._stock_dirty or not getattr(self, "worker", None) \
                or self.worker.busy:
            return False
        if self._current_tab() != 2:
            return False
        self._stock_dirty = False
        self.do_recalc()
        return True

    def _on_unmatched_select(self, _e=None):
        i = self.tbl_unmatched.selected_index()
        if i is None or i >= len(self._unmatched):
            return
        g = self._unmatched[i]
        line = {"raw_name": g["raw_name"], "norm_name": g["norm_name"],
                "mxik": g["mxik"], "barcode": g["barcode"], "marking_code": ""}
        pid, meth, score, sugg = self.engine.match(line)
        cands = list(sugg)
        if pid:
            cands.insert(0, (pid, score, meth))
        self._candidates = cands
        rows = []
        for cid, sc, why in cands:
            p = self.engine.products.get(cid, {})
            left = self.cx.execute(
                "SELECT COALESCE(SUM(CAST(qty_left AS REAL)),0) q FROM stock_lot "
                "WHERE product_id=?", (cid,)).fetchone()["q"]
            rows.append(((p.get("canon_name") or "?")[:80], "%.0f%%" % (sc * 100),
                         why, fmt_qty(left)))
        self.tbl_cand.fill(rows, lambda r: ("ok",) if r[1] >= "88%" else ())
        # Taklif yo'q bo'lsa jadval joy egallamasin - qidiruv natijalari ko'rinsin
        if not rows:
            self.tbl_cand.pack_forget()
            self.lbl_nocand.configure(text="Mos tovar topilmadi - pastdan qidiring.")
            self.lbl_nocand.pack(anchor="w", pady=4)
            self.set_status("Mos tovar topilmadi - qidiruvdan toping")
        else:
            self.lbl_nocand.pack_forget()
            self.tbl_cand.tree.configure(height=min(len(rows), 4))
            self.tbl_cand.pack(fill="x")
            self.set_status("%d ta mos tovar topildi - to'g'risini bosing va \"Saqlash\"" % len(rows))
        # Qidiruvga nomning birinchi mazmunli so'zi (to'liq nom kamdan-kam topiladi)
        words = [w for w in re.split(r"[\s,.()\-]+", g["raw_name"] or "") if len(w) >= 4]
        self.var_search.set(words[0] if words else (g["raw_name"] or "")[:20])
        self._search_products()

    def _search_products(self):
        q = self.var_search.get().strip()
        rows = self.engine.product_choices(q, limit=250)
        self._products = [dict(r) for r in rows]
        self.tbl_prod.fill([((r["canon_name"] or "")[:80], r["mxik"] or "",
                             fmt_qty(r["qty_left"] or 0)) for r in self._products])

    def _run(self, label, job):
        if self.worker.busy:
            messagebox.showinfo("Band", "Oldingi amal hali tugamadi.")
            return
        self._set_busy(True)
        self.worker.start(label, job)

    def _set_busy(self, busy):
        st = "disabled" if busy else "normal"
        for b in (self.btn_add, self.btn_report, self.btn_report_big, self.btn_recalc):
            try:
                b.configure(state=st)
            except tk.TclError:
                pass

    def _on_progress(self, cur, total, msg):
        if total:
            self.progress.configure(mode="determinate", maximum=total, value=cur)
        else:
            self.progress.configure(mode="determinate", maximum=1, value=0)
        if msg:
            self.set_status("%s..." % msg if total else msg)

    def _on_done(self, label, res):
        self.progress.configure(value=0)
        self._set_busy(False)
        try:
            self.cx.close()
        except Exception:
            pass
        self.cx = DB.connect()
        self.engine = M.MatchEngine(self.cx)
        # Import / Qayta hisoblash / Hisobot - uchalasi ham FIFO ni qayta
        # quradi, demak ombor yangi.
        self._stock_dirty = False

        if label == "Import":
            self.queued_files = []
            self._last_results = res.get("results", [])
            ok = sum(1 for r in self._last_results if r[2].startswith("Qabul"))
            bad = sum(1 for r in self._last_results if r[2].startswith(("O'qilmadi", "Fayl ochilmadi")))
            parts = ["Tayyor: %d ta fayl ko'rib chiqildi." % res["files"]]
            if ok:
                parts.append("%d tasi qabul qilindi." % ok)
            if res["dup_files"]:
                parts.append("%d tasi oldin qo'shilgan edi." % res["dup_files"])
            if bad:
                parts.append("%d tasi o'qilmadi - jadvalda sababi yozilgan." % bad)
            self.lbl_import.configure(text="  ".join(parts),
                                      foreground=PALETTE["err"] if bad else PALETTE["ok"],
                                      wraplength=int(900 * self.scale))
            self.set_status("Fayllar qabul qilindi")
            self.nb.select(self.tab_home)
        elif label == "Qayta hisoblash":
            self.set_status("Qayta hisoblandi: %d bog'landi, %d qoldi"
                            % (res["auto"], res["left"]))
        elif label == "Hisobot":
            p, st = os.path.normpath(res["path"]), res["stats"]
            self._last_report = p
            self.btn_open.configure(state="normal")
            self.btn_open_dir.configure(state="normal")
            lines = ["Tayyor! Hisobot saqlandi:", p, ""]
            for y, t in sorted(st["years"].items()):
                lines.append("%d yil:  kirim %s  ·  sotilgan (tannarx) %s  ·  qoldiq %s so'm"
                             % (y, fmt_money(t["in_sum"]).rsplit(".", 1)[0],
                                fmt_money(t["out_sum"]).rsplit(".", 1)[0],
                                fmt_money(t["close_sum"]).rsplit(".", 1)[0]))
            self.lbl_result.configure(text="\n".join(lines), foreground=PALETTE["ok"],
                                      wraplength=int(900 * self.scale))
            self.set_status("Hisobot tayyor - ochilmoqda")
            self._open_last()

        self.refresh_all()

    def _on_error(self, label, tb):
        self.progress.configure(value=0)
        self._set_busy(False)
        if label == "Import":
            self.lbl_import.configure(text="Fayllarni qabul qilishda xato bo'ldi - dasturchiga ayting.")
        elif label == "Hisobot":
            self.lbl_result.configure(text="Hisobot yaratilmadi - xato oynasini dasturchiga yuboring.")
        self.set_status("%s: xato" % label)
        self._show_error("%s amalida xato" % label, tb)

    def _show_error(self, title, tb):
        win = tk.Toplevel(self.root)
        win.title(title)
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        w, h = min(int(820 * self.scale), int(sw * 0.9)), min(int(460 * self.scale),
                                                              int(sh * 0.8))
        win.geometry("%dx%d+%d+%d" % (w, h, (sw - w) // 2, (sh - h) // 3))
        win.minsize(360, 240)

        bar = ttk.Frame(win, padding=(12, 8))
        bar.pack(side="bottom", fill="x")
        ttk.Button(bar, text="Yopish", command=win.destroy).pack(side="right")
        ttk.Button(bar, text="Nusxa olish", style="Ghost.TButton",
                   command=lambda: (self.root.clipboard_clear(),
                                    self.root.clipboard_append(tb))).pack(
            side="right", padx=6)

        ttk.Label(win, text=title, style="H1.TLabel").pack(anchor="w", padx=12,
                                                           pady=(12, 4))
        from tkinter import scrolledtext
        t = scrolledtext.ScrolledText(win, wrap="word", font=self.fonts["mono"])
        t.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        t.insert("1.0", tb)
        t.configure(state="disabled")

    def set_status(self, text):
        self.lbl_status.configure(text=text)

    def _open_last(self):
        if self._last_report:
            self._open_path(self._last_report)

    def _open_path(self, p):
        try:
            if sys.platform.startswith("win"):
                os.startfile(p)             # noqa
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", p])
            else:
                webbrowser.open("file://%s" % p)
        except Exception as e:
            messagebox.showerror("Ochilmadi", str(e))

    def on_close(self):
        try:
            g = self.root.geometry()
            DB.set_setting(self.cx, "window_geometry", g)
            DB.set_setting(self.cx, "folder_kirim", self.var_fak.get())
            DB.set_setting(self.cx, "folder_chiqim", self.var_chq.get())
            self.cx.close()
        except Exception:
            pass
        self.root.destroy()


# ===========================================================================
# Kirish nuqtasi
# ===========================================================================
def main(launcher_info=None):
    # Boshqa joydan ochilgan bo'lsa - doimiy joyga o'rnatib, o'sha nusxa ochiladi
    if ensure_single_install():
        return 0
    root = tk.Tk()
    try:
        App(root, launcher_info)
    except Exception:
        tb = traceback.format_exc()
        try:
            messagebox.showerror("Ishga tushmadi", tb[:3000])
        except Exception:
            sys.stderr.write(tb)
        root.destroy()
        return 1
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
