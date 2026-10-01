@echo off
REM ===========================================================================
REM  BuxgalterHisobot.exe ni yig'ish
REM
REM  Bu buyruq FAQAT launcher o'zgarganda kerak.
REM  Kunlik tuzatishlar uchun: python make_manifest.py && git push
REM ===========================================================================

setlocal
cd /d "%~dp0"

echo.
echo [1/3] manifest.json yangilanmoqda...
python make_manifest.py
if errorlevel 1 goto :error

echo.
echo [2/3] Eski yig'ilma tozalanmoqda...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [3/3] PyInstaller ishga tushmoqda...
REM
REM  MUHIM: launcher.py faqat standart kutubxonani import qiladi.
REM  openpyxl, xlrd, bs4 va h.k. faqat exec() qilingan modullar ichida -
REM  PyInstaller ularni KO'RMAYDI. Shuning uchun har biri qo'lda ko'rsatiladi.
REM
REM  Ro'yxatni KENG oling: core.py keyinchalik yangi kutubxona qo'sha olmaydi,
REM  .exe ichida nima bo'lsa - shu. Kerak bo'lmasa ham qo'shib qo'ying.
REM
pyinstaller --onefile --windowed --clean ^
  --name BuxgalterHisobot ^
  --icon "icon.ico" ^
  --add-data "app;app" ^
  --add-data "manifest.json;." ^
  --add-data "icon.ico;." ^
  --hidden-import openpyxl      --collect-all openpyxl ^
  --hidden-import xlrd          --collect-all xlrd ^
  --hidden-import xlsxwriter    --collect-all xlsxwriter ^
  --hidden-import bs4           --collect-all bs4 ^
  --hidden-import lxml          --collect-all lxml ^
  --hidden-import soupsieve ^
  --hidden-import windnd       --collect-all windnd ^
  --hidden-import tkinterdnd2 ^
  --hidden-import sqlite3 ^
  --hidden-import decimal ^
  --hidden-import difflib ^
  --hidden-import unicodedata ^
  --hidden-import datetime ^
  --hidden-import hashlib ^
  --hidden-import json ^
  --hidden-import csv ^
  --hidden-import re ^
  --hidden-import queue ^
  --hidden-import threading ^
  --hidden-import webbrowser ^
  --hidden-import urllib.request ^
  --hidden-import tkinter ^
  --hidden-import tkinter.ttk ^
  --hidden-import tkinter.filedialog ^
  --hidden-import tkinter.messagebox ^
  --hidden-import tkinter.scrolledtext ^
  --hidden-import tkinter.font ^
  launcher.py
if errorlevel 1 goto :error

echo.
echo ===========================================================
echo  TAYYOR:  dist\BuxgalterHisobot.exe
echo.
echo  Buxgalterga SHU BITTA faylni bering.
echo  Keyingi tuzatishlar uchun .exe qayta kerak EMAS -
echo  faqat: python make_manifest.py  va  git push
echo ===========================================================
goto :end

:error
echo.
echo  XATO: yig'ish tugallanmadi.
exit /b 1

:end
endlocal
