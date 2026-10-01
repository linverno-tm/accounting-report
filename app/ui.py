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
}


def setup_style(root, scale):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    base = max(9, int(round(10 * scale)))
    root.option_add("*Font", ("Segoe UI", base))

    f_base = ("Segoe UI", base)
    f_bold = ("Segoe UI", base, "bold")
    f_h1 = ("Segoe UI", base + 5, "bold")
    f_h2 = ("Segoe UI", base + 1, "bold")
    f_small = ("Segoe UI", max(8, base - 1))
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

    style.configure("TNotebook", background=PALETTE["bg"], borderwidth=0)
    style.configure("TNotebook.Tab", font=f_base,
                    padding=(int(16 * scale), int(8 * scale)))
    style.map("TNotebook.Tab",
              background=[("selected", PALETTE["card"])],
              foreground=[("selected", PALETTE["accent"])])

    rowh = int(round(24 * scale))
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
                                        length=int(150 * self.scale))
        self.progress.pack(side="right", padx=(10, 0))

        self.lbl_counts = ttk.Label(inner, text="", style="Muted.TLabel")
        self.lbl_counts.pack(side="right", padx=(10, 0))

        self.lbl_status = ttk.Label(inner, text="Tayyor", style="Muted.TLabel",
                                    anchor="w")
        self.lbl_status.pack(side="left", fill="x", expand=True)

    def _build_actionbar(self):
        wrap = ttk.Frame(self.root)
        wrap.pack(side="bottom", fill="x")
        ttk.Separator(wrap, orient="horizontal").pack(fill="x")

        bar = FlowBar(wrap, scale=self.scale)
        bar.pack(fill="x", padx=8)
        self.actionbar = bar

        bar.add(ttk.Button(bar, text="Papka tanlash", style="Ghost.TButton",
                           command=self.pick_folder), "left")
        bar.add(ttk.Button(bar, text="Fayl qo'shish", style="Ghost.TButton",
                           command=self.pick_files), "left")
        self.btn_import = ttk.Button(bar, text="Import qilish",
                                     style="Ghost.TButton", command=self.do_import)
        bar.add(self.btn_import, "left")
        self.btn_recalc = ttk.Button(bar, text="Qayta hisoblash",
                                     style="Ghost.TButton", command=self.do_recalc)
        bar.add(self.btn_recalc, "left")
        self.btn_report = ttk.Button(bar, text="Hisobot yaratish",
                                     style="Accent.TButton", command=self.do_report)
        bar.add(self.btn_report, "right")

    def _build_header(self):
        head = ttk.Frame(self.root)
        head.pack(side="top", fill="x", padx=14, pady=(12, 6))

        left = ttk.Frame(head)
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, text=C.APP_TITLE, style="H1.TLabel").pack(anchor="w")

        origins = self.info.get("origins") or {}
        src = "zaxira"
        if origins:
            vals = list(origins.values())
            src = "tarmoq" if all(v == "tarmoq" for v in vals) else \
                  ("kesh" if "kesh" in vals else vals[0])
        self.lbl_sub = ttk.Label(
            left, style="Muted.TLabel",
            text="Versiya %s  ·  kod manbai: %s  ·  baza: %s"
                 % (C.VERSION, src, C.db_path()))
        self.lbl_sub.pack(anchor="w", pady=(2, 0))

    def _build_notebook(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(side="top", fill="both", expand=True, padx=10, pady=(4, 6))

        self.tab_files = self._tab_files()
        self.tab_match = self._tab_match()
        self.tab_stock = self._tab_stock()
        self.tab_report = self._tab_report()
        self.tab_issues = self._tab_issues()
        self.tab_settings = self._tab_settings()

        self.nb.add(self.tab_files, text="  1. Fayllar  ")
        self.nb.add(self.tab_match, text="  2. Moslashtirish  ")
        self.nb.add(self.tab_stock, text="  3. Ombor  ")
        self.nb.add(self.tab_report, text="  4. Hisobot  ")
        self.nb.add(self.tab_issues, text="  5. Xatolar  ")
        self.nb.add(self.tab_settings, text="  6. Sozlamalar  ")
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab)

    # ------------------------------------------------------------------
    # 1. Fayllar
    # ------------------------------------------------------------------
    def _tab_files(self):
        page = ttk.Frame(self.nb, padding=10)

        # --- Tushirish maydoni: bitta fayl ham, 10 ta fayl ham, papka ham ---
        drop = ttk.Frame(page, style="Drop.TFrame")
        drop.pack(fill="x")
        dz = ttk.Frame(drop, style="DropIn.TFrame")
        dz.pack(fill="x", padx=14, pady=12)

        self.lbl_drop = ttk.Label(
            dz, style="H2.TLabel", anchor="center", justify="center",
            background="#eef3fb",
            text="Excel fayllarini shu yerga sudrab tashlang")
        self.lbl_drop.pack(fill="x")
        self.lbl_drop2 = ttk.Label(
            dz, style="CardMuted.TLabel", anchor="center", justify="center",
            background="#eef3fb",
            text="Bitta fayl ham, bir necha fayl ham, butun papka ham bo'ladi - "
                 "hammasi bitta hisobotga qo'shiladi")
        self.lbl_drop2.pack(fill="x", pady=(3, 8))

        dbar = FlowBar(dz, scale=self.scale)
        dbar.pack(fill="x")
        dbar.add(ttk.Button(dbar, text="Fayllarni tanlash", style="Accent.TButton",
                            command=self.pick_files), "left")
        dbar.add(ttk.Button(dbar, text="Papka tanlash", style="Ghost.TButton",
                            command=self.pick_folder), "left")

        # DnD ni oyna va maydonga ulaymiz
        self._dnd = None
        for w in (self.root, drop):
            b = enable_file_drop(w, self.on_files_dropped)
            self._dnd = self._dnd or b
        if not self._dnd:
            self.lbl_drop.configure(
                text="Excel fayllarini tanlang")
            self.lbl_drop2.configure(
                text="Bitta fayl ham, bir necha fayl ham, butun papka ham bo'ladi - "
                     "hammasi bitta hisobotga qo'shiladi")

        # --- Doimiy papkalar (ixtiyoriy) ---
        top, inner = card(page, "Doimiy papkalar (ixtiyoriy)")
        top.pack(fill="x", pady=(10, 0))
        wrapping_label(
            inner,
            "Agar fayllar doim bir joyda tursa, papkani shu yerda belgilab "
            "qo'ying - dastur har ochilganda o'zi skanerlaydi. Fayl turi "
            "kengaytmaga emas, mazmuniga qarab aniqlanadi. Bir marta kiritilgan "
            "fayl ikkinchi marta hisobga olinmaydi."
        ).pack(fill="x", pady=(0, 10))

        self.var_fak = tk.StringVar(value=DB.get_setting(self.cx, "folder_kirim", ""))
        self.var_chq = tk.StringVar(value=DB.get_setting(self.cx, "folder_chiqim", ""))
        for label, var, kind in (("Kirim (FAKTURA)", self.var_fak, "kirim"),
                                 ("Chiqim (CHEKLAR)", self.var_chq, "chiqim")):
            row = ttk.Frame(inner, style="CardIn.TFrame")
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=label, style="Card.TLabel", width=20).pack(side="left")
            ttk.Button(row, text="Tozalash", style="Ghost.TButton",
                       command=lambda v=var, k=kind: self._clear_folder(v, k)
                       ).pack(side="right", padx=(6, 0))
            ttk.Button(row, text="Tanlash", style="Ghost.TButton",
                       command=lambda v=var, k=kind: self._browse_folder(v, k)
                       ).pack(side="right", padx=(8, 0))
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

        # --- Navbat ---
        mid = ttk.Frame(page)
        mid.pack(fill="x", pady=(10, 6))
        ttk.Label(mid, text="Import navbati", style="H1.TLabel").pack(side="left")
        self.lbl_queue = ttk.Label(mid, text="", style="Muted.TLabel")
        self.lbl_queue.pack(side="left", padx=10)
        ttk.Button(mid, text="Navbatni tozalash", style="Ghost.TButton",
                   command=self.clear_queue).pack(side="right")
        ttk.Button(mid, text="Papkalarni skanerlash", style="Ghost.TButton",
                   command=self.scan_folders).pack(side="right", padx=6)
        ttk.Button(mid, text="Satrni o'chirish", style="Ghost.TButton",
                   command=self.remove_queued).pack(side="right", padx=6)

        self.tbl_files = Table(page, [
            ("file", "Fayl", 300, True, "w"),
            ("kind", "Turi", 90, False, "w"),
            ("fmt", "Format", 80, False, "w"),
            ("size", "Hajmi", 80, False, "e"),
            ("state", "Holati", 240, False, "w"),
        ], scale=self.scale, height=9, on_double=self._toggle_kind)
        self.tbl_files.pack(fill="both", expand=True)
        ttk.Label(page, style="Muted.TLabel",
                  text="Maslahat: turi noto'g'ri aniqlangan bo'lsa, satr ustiga "
                       "ikki marta bosing - kirim/chiqim almashadi."
                  ).pack(anchor="w", pady=(4, 0))
        return page

    # ------------------------------------------------------------------
    # 2. Moslashtirish
    # ------------------------------------------------------------------
    def _tab_match(self):
        page = ttk.Frame(self.nb, padding=10)

        top, inner = card(page, "Kirimga bog'lanmagan sotuvlar")
        top.pack(fill="x")
        wrapping_label(
            inner,
            "Kassa mahsulot nomini 63 belgida kesadi, shuning uchun bir qism "
            "sotuv avtomatik bog'lanmaydi. Chapdan nomni tanlang, o'ngdan mos "
            "kirim mahsulotini belgilab Bog'lash tugmasini bosing. Tanlovingiz "
            "eslab qolinadi va bu nom boshqa so'ralmaydi."
        ).pack(fill="x")

        pane = ttk.PanedWindow(page, orient="horizontal")
        pane.pack(fill="both", expand=True, pady=(10, 0))

        left = ttk.Frame(pane)
        ttk.Label(left, text="Bog'lanmagan nomlar", style="Muted.TLabel"
                  ).pack(anchor="w", pady=(0, 4))
        self.tbl_unmatched = Table(left, [
            ("name", "Mahsulot nomi", 240, True, "w"),
            ("n", "Satr", 55, False, "e"),
            ("qty", "Miqdor", 80, False, "e"),
            ("amount", "Summa", 120, False, "e"),
        ], scale=self.scale, height=14, on_select=self._on_unmatched_select)
        self.tbl_unmatched.pack(fill="both", expand=True)

        right = ttk.Frame(pane)
        ttk.Label(right, text="Taklif qilingan mahsulotlar", style="Muted.TLabel"
                  ).pack(anchor="w", pady=(0, 4))
        self.tbl_cand = Table(right, [
            ("name", "Kirim mahsuloti", 230, True, "w"),
            ("score", "Ishonch", 75, False, "e"),
            ("why", "Sabab", 160, False, "w"),
            ("left", "Ombor", 70, False, "e"),
        ], scale=self.scale, height=7, on_double=lambda e: self.do_link("cand"))
        self.tbl_cand.pack(fill="both", expand=True)

        srow = ttk.Frame(right)
        srow.pack(fill="x", pady=(8, 4))
        ttk.Label(srow, text="Qidirish:").pack(side="left")
        self.var_search = tk.StringVar()
        ent = ttk.Entry(srow, textvariable=self.var_search)
        ent.pack(side="left", fill="x", expand=True, padx=6)
        ent.bind("<KeyRelease>", lambda e: self._search_products())
        ttk.Button(srow, text="Bog'lash", style="Accent.TButton",
                   command=lambda: self.do_link()).pack(side="right")

        self.tbl_prod = Table(right, [
            ("name", "Barcha kirim mahsulotlari", 230, True, "w"),
            ("mxik", "MXIK", 130, False, "w"),
            ("left", "Ombor", 70, False, "e"),
        ], scale=self.scale, height=7, on_double=lambda e: self.do_link("prod"))
        self.tbl_prod.pack(fill="both", expand=True)

        # Panellar hech qachon nolgacha siqilmasin
        pane.add(left, weight=3)
        pane.add(right, weight=4)
        return page

    # ------------------------------------------------------------------
    # 3. Ombor
    # ------------------------------------------------------------------
    def _tab_stock(self):
        page = ttk.Frame(self.nb, padding=10)
        row = ttk.Frame(page)
        row.pack(fill="x", pady=(0, 8))
        ttk.Label(row, text="Ombor qoldig'i", style="H1.TLabel").pack(side="left")
        self.var_stock_zero = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="Nol qoldiqni ham ko'rsatish",
                        variable=self.var_stock_zero,
                        command=self.refresh_stock).pack(side="right")

        self.tbl_stock = Table(page, [
            ("name", "Mahsulot", 280, True, "w"),
            ("unit", "Birlik", 90, False, "w"),
            ("inq", "Kirim", 85, False, "e"),
            ("outq", "Sotilgan", 85, False, "e"),
            ("left", "Qoldiq", 85, False, "e"),
            ("cost", "O'rtacha tannarx", 130, False, "e"),
            ("value", "Qoldiq summasi", 140, False, "e"),
        ], scale=self.scale, height=16)
        self.tbl_stock.pack(fill="both", expand=True)
        return page

    # ------------------------------------------------------------------
    # 4. Hisobot
    # ------------------------------------------------------------------
    def _tab_report(self):
        page = ScrollFrame(self.nb)
        pad = ttk.Frame(page.body, padding=10)
        pad.pack(fill="both", expand=True)

        c1, i1 = card(pad, "Hisobot sozlamalari")
        c1.pack(fill="x")

        row = ttk.Frame(i1, style="CardIn.TFrame")
        row.pack(fill="x", pady=3)
        ttk.Label(row, text="Tashkilot / JShDSh", style="Card.TLabel",
                  width=22).pack(side="left")
        self.var_owner = tk.StringVar(
            value=DB.get_setting(self.cx, "owner_name", ""))
        ttk.Entry(row, textvariable=self.var_owner).pack(side="left", fill="x",
                                                         expand=True)

        ttk.Label(i1, text="Qaysi yillar kirsin:", style="Card.TLabel"
                  ).pack(anchor="w", pady=(10, 4))
        self.years_bar = FlowBar(i1, scale=self.scale)
        self.years_bar.pack(fill="x")
        self.year_vars = {}

        row = ttk.Frame(i1, style="CardIn.TFrame")
        row.pack(fill="x", pady=(12, 3))
        ttk.Label(row, text="Saqlash papkasi", style="Card.TLabel",
                  width=22).pack(side="left")
        self.var_outdir = tk.StringVar(
            value=DB.get_setting(self.cx, "out_dir",
                                 os.path.join(os.path.expanduser("~"), "Desktop")))
        ttk.Button(row, text="Tanlash", style="Ghost.TButton",
                   command=self._browse_outdir).pack(side="right", padx=(8, 0))
        ttk.Entry(row, textvariable=self.var_outdir).pack(side="left", fill="x",
                                                          expand=True)

        c2, i2 = card(pad, "Oxirgi natija")
        c2.pack(fill="x", pady=(10, 0))
        self.lbl_result = ttk.Label(
            i2, style="CardMuted.TLabel", justify="left", anchor="w",
            text="Hisobot hali yaratilmadi.\n\n"
                 "Pastdagi \"Hisobot yaratish\" tugmasini bosing.")
        self.lbl_result.pack(fill="x")
        self.btn_open = ttk.Button(i2, text="Faylni ochish", style="Ghost.TButton",
                                   command=self._open_last, state="disabled")
        self.btn_open.pack(anchor="w", pady=(10, 0))

        c3, i3 = card(pad, "Yillar bo'yicha jami")
        c3.pack(fill="both", expand=True, pady=(10, 0))
        self.tbl_years = Table(i3, [
            ("year", "Yil", 70, False, "w"),
            ("inq", "Kirim (tannarx)", 140, True, "e"),
            ("outq", "Chiqim (tannarx)", 140, True, "e"),
            ("close", "Qoldiq (tannarx)", 140, True, "e"),
            ("sale", "Sotuv (QQS bilan)", 140, True, "e"),
            ("profit", "Sof foyda", 130, True, "e"),
        ], scale=self.scale, height=6)
        self.tbl_years.pack(fill="both", expand=True)
        return page

    # ------------------------------------------------------------------
    # 5. Xatolar
    # ------------------------------------------------------------------
    def _tab_issues(self):
        page = ttk.Frame(self.nb, padding=10)
        row = ttk.Frame(page)
        row.pack(fill="x", pady=(0, 8))
        ttk.Label(row, text="Tekshiruv natijalari", style="H1.TLabel").pack(side="left")
        self.var_sev = tk.StringVar(value="hammasi")
        cb = ttk.Combobox(row, textvariable=self.var_sev, state="readonly", width=18,
                          values=["hammasi", C.SEVERITY_ERROR, C.SEVERITY_WARN,
                                  C.SEVERITY_INFO])
        cb.pack(side="right")
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_issues())
        ttk.Label(row, text="Daraja:", style="Muted.TLabel").pack(side="right", padx=6)

        self.tbl_issues = Table(page, [
            ("sev", "Daraja", 110, False, "w"),
            ("year", "Yil", 60, False, "e"),
            ("code", "Turi", 210, False, "w"),
            ("msg", "Tavsif", 400, True, "w"),
        ], scale=self.scale, height=16)
        self.tbl_issues.pack(fill="both", expand=True)
        return page

    # ------------------------------------------------------------------
    # 6. Sozlamalar
    # ------------------------------------------------------------------
    def _tab_settings(self):
        page = ScrollFrame(self.nb)
        pad = ttk.Frame(page.body, padding=10)
        pad.pack(fill="both", expand=True)

        c1, i1 = card(pad, "Ustama stavkasi (yil bo'yicha)")
        c1.pack(fill="x")
        wrapping_label(
            i1,
            "Sotish narxi = tannarx x (1 + ustama) x (1 + QQS). Mavjud "
            "hisobotda ustama yillar bo'yicha har xil edi: 2023-yil 3%, "
            "2024-yil 10%, 2026-yil 5%."
        ).pack(fill="x", pady=(0, 8))

        self.markup_vars = {}
        grid = ttk.Frame(i1, style="CardIn.TFrame")
        grid.pack(fill="x")
        for i, y in enumerate(range(2023, datetime.date.today().year + 2)):
            v = tk.StringVar(value=("%g" % (DB.get_markup(self.cx, y) * 100)))
            self.markup_vars[y] = v
            cell = ttk.Frame(grid, style="CardIn.TFrame")
            cell.grid(row=i // 3, column=i % 3, sticky="w", padx=(0, 22), pady=4)
            ttk.Label(cell, text="%d:" % y, style="Card.TLabel",
                      width=7).pack(side="left")
            ttk.Entry(cell, textvariable=v, width=8).pack(side="left")
            ttk.Label(cell, text="%", style="Card.TLabel").pack(side="left", padx=(4, 0))
        ttk.Button(i1, text="Saqlash", style="Ghost.TButton",
                   command=self.save_markups).pack(anchor="w", pady=(10, 0))

        c2, i2 = card(pad, "Ma'lumotlar bazasi")
        c2.pack(fill="x", pady=(10, 0))
        self.lbl_db = ttk.Label(i2, style="CardMuted.TLabel", justify="left",
                                anchor="w")
        self.lbl_db.pack(fill="x")
        brow = FlowBar(i2, scale=self.scale)
        brow.pack(fill="x", pady=(8, 0))
        brow.add(ttk.Button(brow, text="Zaxira nusxa olish", style="Ghost.TButton",
                            command=self.do_backup), "left")
        brow.add(ttk.Button(brow, text="Papkani ochish", style="Ghost.TButton",
                            command=lambda: self._open_path(C.app_data_dir())), "left")
        brow.add(ttk.Button(brow, text="Hammasini o'chirish", style="Ghost.TButton",
                            command=self.do_reset), "left")

        c4, i4 = card(pad, "Yorliqlar")
        c4.pack(fill="x", pady=(10, 0))
        wrapping_label(
            i4,
            "Ish stolida va Boshlash menyusida yorliq hosil qiladi. Yorliq "
            "tasodifan o'chib ketsa yoki dastur boshqa papkaga ko'chirilsa "
            "shu tugmani bosing."
        ).pack(fill="x", pady=(0, 8))
        ttk.Button(i4, text="Ish stoliga yorliq qo'yish", style="Ghost.TButton",
                   command=self.do_shortcut).pack(anchor="w")

        c3, i3 = card(pad, "Dastur haqida")
        c3.pack(fill="x", pady=(10, 0))
        origins = self.info.get("origins") or {}
        commit = self.info.get("commit")
        txt = ["Core versiya:  %s" % C.VERSION,
               "Launcher:      %s" % self.info.get("launcher_version", "-"),
               "Yangilanish:   %s" % self.info.get("base_url", "-"),
               "Commit:        %s" % (commit[:12] if commit else "branch (5 daq. kechikishi mumkin)"),
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
        self._stale = {1, 2, 3, 4}
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
        if tab == 1:
            self.refresh_unmatched()
            self._search_products()
        elif tab == 2:
            self.refresh_stock()
        elif tab == 3:
            self.refresh_years()
        elif tab == 4:
            self.refresh_issues()

    def refresh_counts(self):
        st = DB.stats(self.cx)
        self.lbl_counts.configure(
            text="Hujjat: %d  ·  Satr: %d  ·  Mahsulot: %d  ·  Bog'lanmagan: %d  ·  Xato: %d"
                 % (st["docs_in"] + st["docs_out"],
                    st["lines_in"] + st["lines_out"],
                    st["products"], st["unmatched"], st["issues"]))

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

    def refresh_queue(self):
        rows = []
        for p, kind in self.queued_files:
            try:
                stt = os.stat(p)
                sha, fmt = self._file_info(p, stt)
                size = stt.st_size
                ex = DB.find_source_file(self.cx, sha)
                if ex is None:
                    state = "Yangi"
                elif F.needs_reparse(self.cx, ex):
                    state = "Yangi (qayta o'qiladi)"
                else:
                    state = "Allaqachon kiritilgan (%s)" % (ex["imported_at"] or "")[:10]
            except OSError:
                size, fmt, state = 0, "?", "Fayl ochilmadi"
            rows.append((os.path.basename(p), kind, fmt,
                         "%.0f KB" % (size / 1024.0), state))

        def tag(r):
            return ("warn",) if r[4].startswith("Allaqachon") else \
                   (("err",) if r[4].startswith("Fayl") else ())

        self.tbl_files.fill(rows, tag)
        n_new = sum(1 for r in rows if r[4].startswith("Yangi"))
        self.lbl_queue.configure(
            text="%d ta fayl, shundan %d ta yangi" % (len(rows), n_new))
        self.btn_import.configure(state="normal" if n_new else "disabled")

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

    def refresh_years(self):
        years = DB.available_years(self.cx)
        for w in list(self.years_bar._left):
            try:
                w.destroy()
            except Exception:
                pass
        self.years_bar._left = []
        self.year_vars = {}
        if not years:
            lbl = ttk.Label(self.years_bar, text="Ma'lumot yo'q - avval import qiling",
                            style="CardMuted.TLabel")
            self.years_bar.add(lbl, "left")
        else:
            for y in years:
                v = tk.BooleanVar(value=True)
                self.year_vars[y] = v
                cb = ttk.Checkbutton(self.years_bar, text=str(y), variable=v)
                self.years_bar.add(cb, "left")

        rows = []
        for y in years:
            t = F.year_totals(F.year_rows(self.cx, y))
            rows.append((y, fmt_money(t["in_sum"]), fmt_money(t["out_sum"]),
                         fmt_money(t["close_sum"]), fmt_money(t["sale_sum"]),
                         fmt_money(t["profit"])))
        self.tbl_years.fill(rows)

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
        self.nb.select(self.tab_files)
        self.refresh_queue()
        dup = len(files) - added
        self.set_status(
            "%d ta fayl qo'shildi%s%s"
            % (added,
               (" (%d tasi navbatda bor edi)" % dup) if dup else "",
               (" - %s" % source) if source else ""))

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
            self.set_status("Papkalarda %d ta fayl, navbatga %d tasi qo'shildi"
                            % (found, added))
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
            for i, (path, kind) in enumerate(files):
                progress(i, len(files), os.path.basename(path))
                try:
                    sid, isnew = DB.add_source_file(cx, path, kind)
                except OSError as e:
                    warns.append("%s: %s" % (os.path.basename(path), e))
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
                    else:
                        dup_files += 1
                    continue
                r = P.parse_any(path, kind)
                warns.extend(r["warnings"])
                a, s, n = F.import_parsed(cx, eng, r, sid)
                DB.update_source_counts(cx, sid, a, n)
                docs += a
                skipped += s
                lines += n
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
                    "left": left, "stock": st, "warns": warns}

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
        years = [y for y, v in self.year_vars.items() if v.get()]
        if not years:
            years = DB.available_years(self.cx)
        if not years:
            messagebox.showwarning(
                "Ma'lumot yo'q",
                "Hisobot yaratish uchun avval fayllarni import qiling.")
            self.nb.select(self.tab_files)
            return

        outdir = self.var_outdir.get().strip() or os.path.expanduser("~")
        owner = self.var_owner.get().strip() or "Ташкилот"
        DB.set_setting(self.cx, "owner_name", owner)
        DB.set_setting(self.cx, "out_dir", outdir)

        path = filedialog.asksaveasfilename(
            title="Hisobotni saqlash", defaultextension=".xlsx",
            initialdir=outdir, initialfile=R.default_filename(years),
            filetypes=[("Excel fayl", "*.xlsx")])
        if not path:
            return

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
        self._stale.update({2, 3})
        self.set_status("Bog'landi: %s  ->  %s" % ((grp["raw_name"] or "")[:40],
                                                   name[:40]))
        self.refresh_unmatched()
        self.refresh_counts()
        if not self._unmatched:
            messagebox.showinfo(
                "Tugadi",
                "Barcha sotuvlar bog'landi.\n\nOmbor avtomatik qayta "
                "hisoblanadi - hisobotda sotilgan tovar qoldiqdan ayriladi.")

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

    def _recalc_if_dirty(self):
        """
        Qo'lda bog'langandan keyin omborni o'zi qayta hisoblaydi.

        Buxgalter "Qayta hisoblash" tugmasini bosishni eslab o'tirmasin:
        "Ombor" yoki "Hisobot" bo'limi ochilganda raqamlar o'zi yangilanadi.
        """
        if not self._stock_dirty or not getattr(self, "worker", None) \
                or self.worker.busy:
            return False
        if self._current_tab() not in (2, 3):
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
        if not self.var_search.get().strip():
            self.var_search.set((g["raw_name"] or "")[:30])
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
        for b in (self.btn_import, self.btn_recalc, self.btn_report):
            b.configure(state="disabled")
        self.worker.start(label, job)

    def _on_progress(self, cur, total, msg):
        if total:
            self.progress.configure(mode="determinate", maximum=total, value=cur)
        else:
            self.progress.configure(mode="determinate", maximum=1, value=0)
        if msg:
            self.set_status("%s..." % msg if total else msg)

    def _on_done(self, label, res):
        self.progress.configure(value=0)
        for b in (self.btn_import, self.btn_recalc, self.btn_report):
            b.configure(state="normal")
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
            msg = ("%d fayl o'qildi (%d tasi allaqachon kiritilgan edi)\n"
                   "%d hujjat, %d satr qo'shildi\n"
                   "Avtomatik bog'landi: %d, qo'lda kerak: %d\n"
                   "FIFO: %d taqsimot, %d yetishmovchilik"
                   % (res["files"], res["dup_files"], res["docs"], res["lines"],
                      res["auto"], res["left"], res["stock"]["allocations"],
                      res["stock"]["shortfalls"]))
            self.set_status("Import tugadi")
            messagebox.showinfo("Import tugadi", msg)
            if res["left"]:
                self.nb.select(self.tab_match)
        elif label == "Qayta hisoblash":
            self.set_status("Qayta hisoblandi: %d bog'landi, %d qoldi"
                            % (res["auto"], res["left"]))
        elif label == "Hisobot":
            p, st = res["path"], res["stats"]
            self._last_report = p
            self.btn_open.configure(state="normal")
            lines = ["Fayl: %s" % p,
                     "Hajmi: %.0f KB" % (os.path.getsize(p) / 1024.0),
                     "Kassa satri: %d   ТХ satri: %d   Xato yozuvi: %d"
                     % (st["kassa_rows"], st["tx_rows"], st["issues"]), ""]
            for y, t in sorted(st["years"].items()):
                lines.append("%d:  kirim %s   chiqim %s   qoldiq %s   foyda %s"
                             % (y, fmt_money(t["in_sum"]), fmt_money(t["out_sum"]),
                                fmt_money(t["close_sum"]), fmt_money(t["profit"])))
            self.lbl_result.configure(text="\n".join(lines))
            self.set_status("Hisobot tayyor")
            if messagebox.askyesno("Tayyor", "Hisobot yaratildi:\n\n%s\n\n"
                                             "Hozir ochamizmi?" % p):
                self._open_last()

        self.refresh_all()

    def _on_error(self, label, tb):
        self.progress.configure(value=0)
        for b in (self.btn_import, self.btn_recalc, self.btn_report):
            b.configure(state="normal")
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
