# -*- coding: utf-8 -*-
"""UI smoke-test: oyna quriladi, turli o'lchamlarda tekshiriladi."""
import sys, os, importlib, tempfile, shutil
sys.path.insert(0, 'app')

# Test bazasini ishlatish uchun config.db_path ni almashtiramiz
import config
TESTDIR = os.path.join(tempfile.gettempdir(), 'bh_uitest')
if os.path.isdir(TESTDIR): shutil.rmtree(TESTDIR, ignore_errors=True)
os.makedirs(TESTDIR, exist_ok=True)
config.app_data_dir = lambda: TESTDIR
config.db_path = lambda: os.path.join(TESTDIR, 'hisobot.db')
sys.modules['bh_config'] = config
for name, mod in [('bh_db','db'),('bh_parsers','parsers'),('bh_matching','matching'),
                  ('bh_fifo','fifo'),('bh_report','report'),('bh_ui','ui')]:
    sys.modules[name] = importlib.import_module(mod)
import bh_ui as U
import tkinter as tk

root = tk.Tk()
app = U.App(root, {'launcher_version':'1.0.0','base_url':'test','origins':{'bh_ui':'zaxira'}})
root.update_idletasks(); root.update()

print('=== OYNA ===')
print('  geometry     :', root.geometry())
print('  minsize      :', root.minsize())
print('  screen       : %dx%d' % (root.winfo_screenwidth(), root.winfo_screenheight()))
print('  DPI scale    : %.2f' % app.scale)

def probe(w, h, label):
    root.geometry('%dx%d' % (w, h))
    root.update_idletasks(); root.update()
    for _ in range(4):
        root.update_idletasks(); root.update()
    ab = app.actionbar
    h = root.winfo_height()          # HAQIQIY balandlik (minsize bilan cheklangan)
    w = root.winfo_width()
    bar_y = ab.winfo_rooty() - root.winfo_rooty()
    bar_h = ab.winfo_height()
    st_y  = app.lbl_status.winfo_rooty() - root.winfo_rooty()
    nb_h  = app.nb.winfo_height()
    # tugmalar joylashuvi
    btns = ab._left + ab._right
    maxx = max((b.winfo_x() + b.winfo_width()) for b in btns)
    rows = len(set(b.winfo_y() for b in btns))
    ok_bar   = (bar_y + bar_h) <= h + 2
    ok_stat  = (st_y + app.lbl_status.winfo_height()) <= h + 2
    ok_width = maxx <= ab.winfo_width() + 2
    print('  %-22s oyna=%dx%d  panel_y=%-4d balandlik=%-3d qator=%d  '
          'eng_o\'ng=%-4d(panel %d)  notebook_h=%-4d  %s'
          % (label, w, h, bar_y, bar_h, rows, maxx, ab.winfo_width(), nb_h,
             'OK' if (ok_bar and ok_stat and ok_width) else 'XATO'))
    return ok_bar and ok_stat and ok_width

print()
print('=== TURLI EKRAN O\'LCHAMLARIDA ===')
res = []
for w,h,lbl in [(1240,780,'katta (1240x780)'),(1024,700,'noutbuk (1024x700)'),
                (900,620,'kichik (900x620)'),(760,560,'juda kichik (760x560)'),
                (640,520,'minimal (640x520)'),(520,480,'ekstremal (520x480)')]:
    res.append(probe(w,h,lbl))

print()
print('=== HAR BIR VARAQ ===')
for i in range(len(app.nb.tabs())):
    app.nb.select(i)
    root.update_idletasks(); root.update()
    name = app.nb.tab(i,'text').strip()
    print('  %-22s ochildi, notebook balandligi=%d' % (name, app.nb.winfo_height()))

print()
print('NATIJA:', 'HAMMASI OK' if all(res) else 'MUAMMO BOR')
root.destroy()
