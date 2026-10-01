# Buxgalteriya hisoboti generatori

Kirim (soliq.uz hisob-fakturalari) va chiqim (kassa cheklari) fayllarini
o'qib, "КАМЕРАЛ ТЕКШИРУВЛАР" ko'rinishidagi Excel hisobotini yig'adi.

## Buxgalterga yetkazish

1. `build.bat` ni bir marta ishga tushiring -> `dist\BuxgalterHisobot.exe`
2. Shu **bitta faylni** buxgalterga bering (Telegram, flesh, tarmoq papkasi - farqi yo'q)
3. U faylni istalgan joyga qo'yib ishga tushiradi. O'rnatish, Python, admin huquqi kerak emas.

Buxgalterda saqlanadigan narsalar:
```
%LOCALAPPDATA%\BuxgalterHisobot    hisobot.db          <- baza (butun tarix shu yerda)
    backup\             <- avtomatik zaxira nusxalar
    code\               <- GitHub'dan yuklangan kod keshi
```

`.exe` ni yangisiga almashtirsangiz ham baza joyida qoladi.

## Tuzilishi

```
launcher.py        -> .exe ga kompilyatsiya qilinadi. DEYARLI O'ZGARMAYDI.
manifest.json      -> modullar ro'yxati + sha256 (make_manifest.py yozadi)
app/
  config.py        (bh_config)   qoidalar, lug'atlar, raqam/sana o'qish
  db.py            (bh_db)       SQLite sxema, migratsiya
  parsers.py       (bh_parsers)  3 xil manba formati
  matching.py      (bh_matching) sotuvni kirimga bog'lash
  fifo.py          (bh_fifo)     FIFO ombor, tekshiruv
  report.py        (bh_report)   Excel yozish
  ui.py            (bh_ui)       tkinter interfeys, main()
```

## Yangilanish qanday ishlaydi

Buxgalterda faqat `.exe` turadi. Har ochilganda:

1. GitHub raw -> `manifest.json`
2. Har modul yuklanadi, **sha256 tekshiriladi**, keshga yoziladi
3. `exec()` bilan modulga aylantiriladi (bog'liqlik tartibida)
4. `bh_ui.main()` chaqiriladi

Tarmoq yo'q bo'lsa: kesh -> `.exe` ichidagi zaxira. Dastur hech qachon
butunlay ishlamay qolmaydi.

### Nega commit sha orqali

`raw.githubusercontent.com/.../main/...` manzili GitHub CDN'ida **~5 daqiqa
keshlanadi** - `git push` dan keyin darhol ochilsa eski kod keladi.
`Cache-Control: no-cache` ham, `?t=123` parametri ham bu keshni kesmaydi
(ikkalasi ham sinab ko'rilgan).

Shuning uchun launcher avval eng so'nggi commit sha'ni aniqlaydi va
**o'zgarmas** `raw.githubusercontent.com/.../<sha>/...` manzilidan yuklaydi -
har commit uchun yangi URL, demak hech qachon eski bo'lmaydi.

Sha qanday aniqlanadi:
1. GitHub API (`/commits/main`) - keshlanmaydi, lekin soatiga 60 so'rov
2. Atom feed (`/commits/main.atom`) - chegara yo'q, qisqa kesh bo'lishi mumkin
3. Ikkalasi ham ishlamasa - oddiy `main` manzili (5 daqiqagacha kechikish)

Amalda: buxgalter dasturni yopib-ochsa yangi kod keladi.

### Tuzatish chiqarish

```
python make_manifest.py
git add -A && git commit -m "tuzatish" && git push
```

Buxgalter dasturni yopib-ochsa - yangi kod. `.exe` qayta kerak emas.

### `.exe` qachon qayta yig'iladi

Faqat `launcher.py` o'zgarganda yoki **yangi kutubxona** kerak bo'lganda.
`app/` modullari `.exe` ichida bo'lmagan kutubxonani import qila olmaydi -
shuning uchun `build.bat` dagi `--hidden-import` ro'yxati keng olingan.

```
build.bat
```

## Fayl qo'shishning 4 yo'li

Hammasi bir xil natija beradi - bitta navbat, bitta hisobot:

1. **Sudrab tashlash** - bitta fayl, 10 ta fayl yoki butun papkani oyna
   ustiga tashlang (`windnd` orqali; kutubxona bo'lmasa jimgina o'chadi)
2. **Fayllarni tanlash** - Ctrl/Shift bilan bir nechtasini belgilash mumkin
3. **Papka tanlash** - papkadagi hammasi
4. **Doimiy papkalar** - bir marta belgilansa, har ochilganda avtomatik skanerlanadi

Fayl turi (kirim/chiqim) mazmuniga qarab aniqlanadi. Noto'g'ri chiqsa -
navbat satriga **ikki marta bosing**, almashadi.

Bir marta import qilingan fayl (sha256 bo'yicha) ikkinchi marta hisobga
olinmaydi - xohlagancha qayta tashlash mumkin.

## Ishlab chiqish rejimi

```
python run_dev.py        # GitHub'siz, to'g'ridan-to'g'ri app/ dan
python test_pipeline.py  # parser -> FIFO -> hisobot zanjiri
python test_ui.py        # oyna turli o'lchamlarda tekshiriladi
python test_flow.py      # tugmalar paneli torayganda o'raladimi
python test_drop.py      # bitta / ko'p / papka / aralash fayl qo'shish
python test_launcher.py  # zaxira zanjiri (tarmoq yo'q holati)
```

## Manba formatlari haqida bilib qo'yish kerak bo'lgan narsalar

Bular haqiqiy fayllar ustida o'lchandi - kodda ham izohlangan:

| Narsa | Haqiqat |
|---|---|
| `FAKTURA/*.xls` | Aslida **HTML**, `.xls` emas. Ustunlar soni 10/11/12 - o'zgarib turadi. Faqat sarlavha matni bo'yicha o'qiladi. |
| Chekdagi `Нархи` | **Narx emas, satr summasi** (QQS ichida). 302 chekda tekshirilgan. Birlik narxi = `Нархи / Миқдори`. |
| Chekdagi `Жами нақд/карта` | Chek jami, lekin **har satrda takrorlanadi**. Satr bo'yicha yig'ish tushumni 2.26 barobar shishiradi. |
| Fakturadagi `Нарҳ` | Birlik narxi, **QQS'siz**. Chekdagidan teskari konvensiya. |
| Kassa mahsulot nomi | **63 belgida kesiladi**. Shuning uchun prefiks bo'yicha moslashtiriladi. |
| MXIK kodi | Kalit bo'la olmaydi: `08517001001000000` -> 21 xil telefon. |
| Ustama | Yillik emas, satr bo'yicha: 3% / 5% / 10% aralash uchraydi. |
| `checks-info` varaqi | `<dimension>` yozuvi noto'g'ri bo'lishi mumkin - `reset_dimensions()` siz oxirgi oylar jimgina tushib qoladi. Barcha varaqlar o'qiladi. |
| Kassa nomi `MEBEL` | Fakturada umuman yo'q tovar. Bog'lanmagan sotuv ham ТХ ga qizil satr bo'lib tushadi - aks holda butun oy yo'qoladi. |
| Bir tovar - ikki kartochka | Faktura nomi ham kesiladi (`...240` / `...240 Вт`). MXIK bir xil va nom prefiks bo'lsa - bitta tovar, ombor birga yuritiladi. |

## Eski hisobotdagi nuqsonlar (bu ilova tuzatadi)

- Yillararo qoldiq o'tmagan (`давр бошига қолдиқ` 4 yilda ham nol)
- Manfiy qoldiq (2026 jami: -6 372.92 dona)
- 2026 da 7 satrda manfiy tannarx
- 2025 varag'ida 2026-yil sanali 15 satr
