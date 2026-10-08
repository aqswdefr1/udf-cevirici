#!/usr/bin/env python3
"""
UDF Çevirici — masaüstü uygulaması.

Word (.docx) belgesini, biçimini koruyarak UYAP Doküman Editörü'nün UDF biçimine;
UDF belgesini de Word'e çevirir. Yönü dosyanın uzantısı belirler. .doc, .rtf, .odt ve .pages
belgeleri önce bilgisayardaki Word'e (Pages belgesi için Pages'e) Word biçimine çevirtilir. Hiçbir dış programa
bağlı değildir; belgeler bilgisayardan çıkmaz. İnternet yalnızca kullanıcı güncelleme
denetimi düğmesine bastığında kullanılır.

Akış:  belgeleri ekle → çevir → (her çıktı kaynağıyla karşılaştırılıp doğrulanır) → rapor / önizleme
"""
import json
import os
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
import zipfile
from tkinter import filedialog, messagebox, ttk

if getattr(sys, "frozen", False):                   # PyInstaller paketi
    sys.path.insert(0, os.path.dirname(os.path.abspath(sys.executable)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import md_udf
import ofis_docx
import udf_docx
import udf_onizle
from docx_udf import DonusumHatasi, cevir
from udf_dogrula import dogrula, dogrula_docx

# Sürükle-bırak Tkinter'da yerleşik değil; paket yoksa düğmeyle devam edilir.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TEMEL_PENCERE, SURUKLENEBILIR = TkinterDnD.Tk, True
except Exception:
    TEMEL_PENCERE, SURUKLENEBILIR = tk.Tk, False

UYGULAMA_ADI = "UDF Çevirici"
SURUM = "1.2"
YAZAR = "Av. Arb. Mevlana İbrahim Asım Bilir"
YAZAR_EK = "av.ibrahimbilir@gmail.com"
TELIF = "© 2026 Av. Arb. Mevlana İbrahim Asım Bilir"
LISANS = "Ücretsiz kullanılabilir ve dağıtılabilir; satılamaz."

# Güncelleme denetimi YALNIZCA kullanıcı düğmeye bastığında yapılır; program
# kendiliğinden internete çıkmaz ve hiçbir belge gönderilmez. Depo herkese
# açıldığında adı buraya yazılır; boşken düğme hiç gösterilmez.
DEPO = "miasimbilir/udf-cevirici"
SURUM_API = f"https://api.github.com/repos/{DEPO}/releases/latest"
SURUM_SAYFA = f"https://github.com/{DEPO}/releases/latest"

SAYFA_NO = {"Word'deki gibi": "otomatik", "Ekle": "var", "Ekleme": "yok"}
KABUL = (".docx", ".udf", ".md", ".markdown") + ofis_docx.UZANTILAR      # bırakılabilen belgeler
DOLGU = {"Renk bandı": "bant", "Yalnız yazı": "yazi", "Yok": "yok"}


def surum_sayilari(etiket):
    """'v1.2' → (1, 2). Karşılaştırma için sayıya çevirir."""
    parcalar = []
    for p in str(etiket).lstrip("vVsS").split("."):
        rakam = "".join(c for c in p if c.isdigit())
        parcalar.append(int(rakam) if rakam else 0)
    return tuple(parcalar) or (0,)


def son_surumu_sor():
    """Depodaki en son sürüm etiketini döndürür. Ağ hatasında istisna atar."""
    istek = urllib.request.Request(SURUM_API, headers={
        "User-Agent": f"UDF-Cevirici/{SURUM}", "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(istek, timeout=8) as yanit:
        return json.load(yanit).get("tag_name") or ""


# Yüksek Çözünürlüklü Ekran (HiDPI / Retina) Desteği
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

IKON_DIZINI = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ikon")
ICO_YOLU = os.path.join(IKON_DIZINI, "uygulama.ico")
PNG_YOLU = os.path.join(IKON_DIZINI, "ikon-1024.png")


def sistem_fontu():
    if sys.platform == "win32":
        return "Segoe UI"
    elif sys.platform == "darwin":
        return ".AppleSystemUIFont"
    return "DejaVu Sans"


def sistem_mono_fontu():
    if sys.platform == "darwin":
        return "Menlo"
    return "Consolas"


def sistem_koyu_mu():
    """İşletim sisteminin koyu temada olup olmadığını anlar."""
    if sys.platform == "win32":
        try:
            import winreg
            anahtar = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            )
            deger, _ = winreg.QueryValueEx(anahtar, "AppsUseLightTheme")
            return deger == 0
        except Exception:
            pass
    elif sys.platform == "darwin":
        try:
            cikti = subprocess.check_output(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                stderr=subprocess.DEVNULL
            ).decode().strip()
            return "Dark" in cikti
        except Exception:
            pass
    return False


def pencerelere_koyu_baslik_uygula(pencere, koyu_mu):
    """Windows 10/11'de pencere başlık çubuğunun rengini temaya göre ayarlar."""
    if sys.platform == "win32":
        try:
            import ctypes
            pencere.update_idletasks()
            hwnd = ctypes.windll.user32.GetAncestor(pencere.winfo_id(), 2)
            if hwnd:
                deger = ctypes.c_int(1 if koyu_mu else 0)
                # Windows 11 ve Windows 10 2004+ için DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                res = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, 20, ctypes.byref(deger), ctypes.sizeof(deger)
                )
                if res != 0:
                    # Windows 10 1903/1909 için DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, 19, ctypes.byref(deger), ctypes.sizeof(deger)
                    )
        except Exception:
            pass


def pencereyi_ortala(pencere, ana=None, w=None, h=None):
    """Pencereyi ana pencerenin veya ekranın ortasına konumlandırır."""
    pencere.update_idletasks()
    pw = w or pencere.winfo_reqwidth()
    ph = h or pencere.winfo_reqheight()
    if ana and ana.winfo_ismapped():
        ax = ana.winfo_rootx()
        ay = ana.winfo_rooty()
        aw = ana.winfo_width()
        ah = ana.winfo_height()
        x = ax + max(0, (aw - pw) // 2)
        y = ay + max(0, (ah - ph) // 2)
    else:
        sw = pencere.winfo_screenwidth()
        sh = pencere.winfo_screenheight()
        x = max(0, (sw - pw) // 2)
        y = max(0, (sh - ph) // 2)
    pencere.geometry(f"{pw}x{ph}+{x}+{y}")


PALETLER = {
    "acik": {
        "ad": "Açık",
        "bg_pencere": "#f8fafc",      # Slate 50
        "bg_kart": "#ffffff",         # Beyaz
        "bg_kart_alt": "#f1f5f9",     # Slate 100
        "cizgi": "#e2e8f0",           # Slate 200
        "cizgi_odak": "#3b82f6",      # Mavi 500
        "yazi": "#0f172a",            # Slate 900
        "yazi_ikincil": "#475569",    # Slate 600
        "soluk": "#64748b",           # Slate 500
        "soluk_acik": "#94a3b8",      # Slate 400
        "birincil": "#2563eb",        # Mavi 600
        "birincil_hover": "#1d4ed8",  # Mavi 700
        "birincil_yazi": "#ffffff",
        "ikincil_bg": "#f1f5f9",      # Buton bg
        "ikincil_cizgi": "#cbd5e1",
        "ikincil_hover": "#e2e8f0",
        "ikincil_yazi": "#1e293b",
        "secim": "#dbeafe",           # Mavi 100
        "secim_yazi": "#1e3a8a",
        "drop_bg": "#ffffff",
        "drop_hover": "#eff6ff",      # Mavi 50
        "drop_cizgi": "#cbd5e1",
        "drop_cizgi_hover": "#3b82f6",
        # Durum renkleri
        "yesil": "#16a34a",
        "yesil_bg": "#f0fdf4",
        "yesil_cizgi": "#86efac",
        "yesil_yazi": "#14532d",
        "sari": "#d97706",
        "sari_bg": "#fffbeb",
        "sari_cizgi": "#fde68a",
        "sari_yazi": "#78350f",
        "kirmizi": "#dc2626",
        "kirmizi_bg": "#fef2f2",
        "kirmizi_cizgi": "#fca5a5",
        "kirmizi_yazi": "#7f1d1d",
        # Geriye uyumluluk
        "kagit": "#ffffff",
    },
    "koyu": {
        "ad": "Koyu",
        "bg_pencere": "#0f172a",      # Slate 900
        "bg_kart": "#1e293b",         # Slate 800
        "bg_kart_alt": "#172033",
        "cizgi": "#334155",           # Slate 700
        "cizgi_odak": "#60a5fa",      # Mavi 400
        "yazi": "#f8fafc",            # Slate 50
        "yazi_ikincil": "#cbd5e1",    # Slate 300
        "soluk": "#94a3b8",           # Slate 400
        "soluk_acik": "#64748b",      # Slate 500
        "birincil": "#3b82f6",        # Mavi 500
        "birincil_hover": "#2563eb",  # Mavi 600
        "birincil_yazi": "#ffffff",
        "ikincil_bg": "#283548",
        "ikincil_cizgi": "#475569",
        "ikincil_hover": "#334155",
        "ikincil_yazi": "#f1f5f9",
        "secim": "#1e3a8a",           # Mavi 900
        "secim_yazi": "#dbeafe",
        "drop_bg": "#1e293b",
        "drop_hover": "#172554",      # Mavi 950
        "drop_cizgi": "#475569",
        "drop_cizgi_hover": "#60a5fa",
        # Durum renkleri
        "yesil": "#22c55e",
        "yesil_bg": "#052e16",
        "yesil_cizgi": "#15803d",
        "yesil_yazi": "#bbf7d0",
        "sari": "#f59e0b",
        "sari_bg": "#451a03",
        "sari_cizgi": "#b45309",
        "sari_yazi": "#fde68a",
        "kirmizi": "#ef4444",
        "kirmizi_bg": "#450a0a",
        "kirmizi_cizgi": "#b91c1c",
        "kirmizi_yazi": "#fecaca",
        # Geriye uyumluluk
        "kagit": "#1e293b",
    }
}
ACIK_TEMA = PALETLER["acik"]
KOYU_TEMA = PALETLER["koyu"]


def tema_sec(pencere=None):
    """Arka planın parlaklığına göre okunur renk kümesi."""
    if sistem_koyu_mu():
        return KOYU_TEMA
    if pencere:
        try:
            r, g, b = pencere.winfo_rgb(ttk.Style().lookup("TFrame", "background")
                                        or pencere.cget("background"))
            return KOYU_TEMA if (r + g + b) / 3 < 32768 else ACIK_TEMA
        except Exception:
            pass
    return ACIK_TEMA


# --------------------------------------------------------------------------
# Çevirme işi (arayüzden bağımsız; sınama da bunu kullanır)
# --------------------------------------------------------------------------
def bos_ad(klasor, ad, uzanti=".docx"):
    """Var olan bir dosyanın üzerine yazmayacak ad: 'A.docx' doluysa "A (UDF'den).docx", o da doluysa 2, 3…"""
    aday = os.path.join(klasor, ad + uzanti)
    n = 1
    while os.path.exists(aday):
        aday = os.path.join(klasor, f"{ad} (UDF'den{'' if n == 1 else ' ' + str(n)}){uzanti}")
        n += 1
    return aday


def belge_cevir(kaynak, hedef_klasor=None, sayfa_no="otomatik", dolgu="bant"):
    """Bir belgeyi çevirir ve DOĞRULAR. Doğrulama geçmezse çıktı bırakılmaz.

    Yönü uzantı belirler: .docx → .udf, .udf → .docx. Word'e çevirirken var olan bir Word
    belgesinin üzerine ASLA yazılmaz (asıl belge çoğu zaman UDF'nin yanında durur).
    Döner: {'kaynak', 'cikti', 'yon', 'tamam', 'ozet', 'uyarilar', 'dogrulama', 'hata'}
    """
    ad, uzanti = os.path.splitext(os.path.basename(kaynak.rstrip("/\\")))
    klasor = hedef_klasor or os.path.dirname(os.path.abspath(kaynak.rstrip("/\\")))
    ters = uzanti.lower() == ".udf"
    is_md = uzanti.lower() in (".md", ".markdown")
    ofis = ofis_docx.destekli(kaynak)
    s = {"kaynak": kaynak, "cikti": "", "yon": "udf>word" if ters else ("md>udf" if is_md else "word>udf"), "tamam": False,
         "ozet": None, "uyarilar": [], "dogrulama": [], "hata": ""}
    if uzanti.lower() not in (".docx", ".udf", ".md", ".markdown") and not ofis:
        s["hata"] = "yalnız Word (.docx, .doc, .rtf, .odt), Pages, Markdown (.md) ve UDF belgeleri çevrilir."
        return s
    cikti = bos_ad(klasor, ad) if ters else os.path.join(klasor, ad + ".udf")
    s["cikti"] = cikti
    gecici = cikti + ".gecici"
    ara = None                                              # .doc/.rtf/.odt/.pages -> geçici .docx
    try:
        if ters:
            s["ozet"] = udf_docx.cevir(kaynak, gecici)
            tamam, satirlar = dogrula_docx(kaynak, gecici)
        elif is_md:
            s["ozet"] = md_udf.cevir(kaynak, gecici, sayfa_no=sayfa_no)
            tamam, satirlar = dogrula(kaynak, gecici)
        else:
            docx, on_uyari = kaynak, []
            if ofis:
                ara = os.path.join(tempfile.mkdtemp(prefix="udf_cevirici_"), "belge.docx")
                program = ofis_docx.docx_yap(kaynak, ara)
                docx = ara
                on_uyari = [f"belge önce {program} ile Word biçimine çevrildi; UDF ve doğrulama bu hâline göre "
                            "yapıldı"]
            s["ozet"] = cevir(docx, gecici, sayfa_no, dolgu)
            s["ozet"]["uyarilar"] = on_uyari + s["ozet"]["uyarilar"]
            tamam, satirlar = dogrula(docx, gecici)
        s["uyarilar"] = s["ozet"]["uyarilar"]
        s["dogrulama"] = satirlar
        if not tamam:
            s["hata"] = "çıktı doğrulamadan geçmedi; güvenli olmadığı için dosya yazılmadı."
            return s
        if ters and os.path.exists(cikti):                  # araya başka bir dosya girdiyse yine ezme
            cikti = s["cikti"] = bos_ad(klasor, ad)
        os.replace(gecici, cikti)
        s["tamam"] = True
    except (DonusumHatasi, md_udf.DonusumHatasi, udf_docx.DonusumHatasi, ofis_docx.OfisHatasi) as e:
        s["hata"] = str(e)
    except OSError as e:
        s["hata"] = f"dosya yazılamadı: {e}"
    except Exception as e:                                  # beklenmeyen: programı düşürme
        s["hata"] = f"beklenmeyen hata ({type(e).__name__}): {e}"
    finally:
        if os.path.exists(gecici):
            try:
                os.remove(gecici)
            except OSError:
                pass
        if ara:
            import shutil
            shutil.rmtree(os.path.dirname(ara), ignore_errors=True)
    return s


# Çekirdeğin teknik uyarıları → kullanıcının anlayacağı kısa cümleler (sonuç kutusunda gösterilir;
# teknik metin Rapor'da kalır). Eşleşme alt dizgiyle yapılır; eşleşmeyen uyarı olduğu gibi gösterilir.
SADE_UYARI = (
    ("dipnotlar belge sonuna", "Dipnotlar sayfa altında değil, belgenin sonunda yer alıyor (UDF'de sayfa altı dipnot yok)."),
    ("içindekiler tablosundaki", "İçindekiler tablosundaki sayfa numaraları Word'deki hâliyle yazıldı; Editör'de sayfalar kayarsa numaralar tutmayabilir."),
    ("hücre dolgusu UDF", "Tablo hücrelerinin dolgu rengi UDF'de yok; renk, yazının arkasına bant olarak verildi."),
    ("hücre gölgelendirmesi", "Tablo hücrelerinin dolgu rengi taşınmadı."),
    ("dikey birleştir", "Dikey birleştirilmiş bazı tablo hücreleri taşınamadı; alttaki hücre boş kaldı."),
    ("korundu; devam sayfalarının", "İlk sayfanın üstbilgisi/altbilgisi korundu; devam sayfalarındaki farklı olan taşınmadı (Editör tek üstbilgi tutuyor)."),
    ("metin kaydırmalı görsel", "Yanından metin akan görsel iki sütunlu tabloyla taklit edildi; yeri Word'dekinden biraz farklı olabilir."),
    ("serbest konumlu görsel", "Serbest konumlu ya da yanından metin akan görsel satır içine alındı; yeri Word'dekinden farklı olabilir."),
    ("metin kutusu", "Metin kutusu düz paragrafa çevrildi; kutunun yeri ve çerçevesi taşınmadı."),
    ("şekil/grafik", "Şekil, grafik ya da SmartArt taşınmadı."),
    ("desteklenmeyen görsel", "Bir görsel (EMF/WMF/PDF gibi) çizilemediği için taşınamadı."),
    ("çevrilemedi, atlandı", "Bir görsel PNG'ye çevrilemediği için taşınamadı."),
    ("VML görsel", "Eski tip bir görsel taşınamadı."),
    ("emoji", "Editör'ün açamadığı karakterler (emoji gibi) □ ile değiştirildi."),
    ("simgenin Unicode", "Bazı özel simgeler □ olarak yazıldı."),
    ("sayfa boyutu Letter", "Sayfa boyutu Letter (Amerikan); A4 istiyorsanız Word'de sayfa boyutunu A4 yapıp yeniden çevirin."),
    ("tanıdığı boyutlardan değil", "Bu sayfa boyutu Editör'de yok; belge A4 yazıldı, satır sonları Word'dekinden farklı olabilir."),
    ("hücre içindeki tablo", "Tablo içindeki tablo düz paragraflara çevrildi."),
    ("hücresindeki sayfa sonu", "Tablo içindeki sayfa sonu atlandı."),
    ("sütun sayısından fazla hücre", "Bir tablo satırında fazla hücre vardı; fazlası atıldı."),
    ("harfli listede Editör Türk", "Harfli listede Editör Türk alfabesiyle sayar: Word'deki d, e… bentleri Editör'de "
                                   "ç, d… görünür. Metinde bent atıfı varsa kontrol edin."),
    ("çok basamaklı numaralandırma", "\"1.1.\" gibi çok basamaklı numaralar Editör'de yok; yalnız son basamak numaralandı."),
    ("ilk sayfanınki kullanıldı, devam sayfalarındaki", "Antet sayfa zemini olarak yazıldı. İlk sayfanın anteti kullanıldı; "
                                                        "devam sayfalarındaki farklı antet taşınmadı."),
    ("Editör'de her sayfada görünür", "Antet sayfa zemini olarak yazıldı. Word'de yalnız bazı sayfalardaydı; "
                                      "Editör'de her sayfada görünür."),
    ("UDF'de sayfa arka planı olarak yazıldı", "Antet, UDF'de sayfa zemini olarak yazıldı (Editör'de her sayfada görünür)."),
    ("ile Word biçimine çevrildi", None),                   # aşağıda programın adıyla doldurulur
    # --- UDF → Word ---
    ("e-imzalı", "Bu UDF e-imzalıydı; Word belgesi imza taşımaz, yalnızca çalışma kopyasıdır."),
    ("harf sıralı liste", "Harfli liste etiketleri (a, b, c, ç…) Editör'deki hâliyle düz metin yazıldı; "
                          "Word'de otomatik numara değildir."),
    ("antet (arka plan görseli) Word", "Antet, Word'de üstbilgiye bağlı arka plan görseli olarak yerleştirildi."),
    ("antet (arka plan görseli) okunamadı", "Antet görseli okunamadığı için taşınamadı."),
    ("form alanları", "Form alanları düz metne çevrildi."),
    ("barkod", "Barkod taşınmadı."),
    ("tanınmayan öğe", "Editör'e özgü bazı öğelerin metni korundu, biçimi taşınmadı."),
    ("sayfa aralığı", "Üstbilgi/altbilgi yalnız belirli sayfalar için ayarlanmıştı; Word'de bütün sayfalara konuldu."),
    ("birden çok", "Belgede birden çok üstbilgi/altbilgi vardı; Editör'ün yaptığı gibi sonuncusu alındı."),
    ("okunamayan bir görsel", "Bir görsel okunamadığı için atlandı."),
    ("tanınmayan kâğıt kodu", "Kâğıt boyutu tanınmadı; A4 kullanıldı."),
    ("tanınmayan tablo kenarlık", "Bir tablonun kenarlık türü tanınmadı; tam kenarlık kullanıldı."),
)


def sade_uyarilar(sonuclar, en_cok=4):
    """Bütün belgelerin uyarılarını sadeleştirir, yineleneni atar. Döner: (gösterilecekler, kalan sayısı)."""
    gorulen, liste = set(), []
    for r in sonuclar:
        for u in r["uyarilar"]:
            sade = next((m for a, m in SADE_UYARI if a in u), u[:1].upper() + u[1:])
            if sade is None:                                # "belge önce Microsoft Word ile Word biçimine…"
                program = u.split("belge önce ", 1)[-1].split(" ile ", 1)[0]
                sade = f"Belge önce {program} ile Word biçimine çevrildi, sonra UDF yapıldı."
            if sade not in gorulen:
                gorulen.add(sade)
                liste.append(sade)
    return liste[:en_cok], max(0, len(liste) - en_cok)


def rapor_metni(sonuclar):
    s = ["UDF ÇEVİRİCİ — ÇEVİRİ RAPORU", ""]
    for i, r in enumerate(sonuclar, 1):
        yon_str = 'UDF → Word' if r['yon'] == 'udf>word' else ('Markdown → UDF' if r['yon'] == 'md>udf' else 'Word → UDF')
        s.append(f"{i}. {os.path.basename(r['kaynak'])}   ({yon_str})")
        if r["tamam"]:
            o = r["ozet"]
            s.append(f"   ✓ Çevrildi → {r['cikti']}")
            ek = [f"{o['paragraf']} paragraf", f"{o['tablo']} tablo"]
            if o.get("gorsel"):
                ek.append(f"{o['gorsel']} görsel")
            if o.get("dikey_birlesik"):
                ek.append("dikey birleşik hücreler Word'ün kendi birleştirmesine çevrildi")
            if o["sayfa_sonu"]:
                ek.append(f"{o['sayfa_sonu']} sayfa sonu")
            if o["ustbilgi"]:
                ek.append("üstbilgi")
            if o["altbilgi"]:
                ek.append("altbilgi")
            s.append("   " + ", ".join(ek))
        else:
            s.append(f"   ✕ Çevrilemedi: {r['hata']}")
        for u in r["uyarilar"]:
            s.append(f"   ! {u}")
        if r["dogrulama"]:
            s.append("   Doğrulama:")
            s += [f"     {d}" for d in r["dogrulama"]]
        s.append("")
    if any(r["yon"] in ("word>udf", "md>udf") for r in sonuclar):
        s += ["UDF'yi UYAP'a yüklemeden önce UYAP Doküman Editörü'nde açıp bir kez gözle",
              "kontrol etmeniz önerilir."]
    if any(r["yon"] == "udf>word" for r in sonuclar):
        s += ["Word'e çevrilen belgede satır ve sayfa sonları Editör'dekinden farklı düşebilir;",
              "görseller UDF'deki kalitesiyle, yeniden sıkıştırılmadan aktarılır."]
    return "\n".join(s)


def klasorde_goster(yol):
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", "-R", yol])
        elif os.name == "nt":
            subprocess.run(["explorer", "/select,", os.path.normpath(yol)])
        else:
            subprocess.run(["xdg-open", os.path.dirname(yol)])
    except Exception:
        pass


def dosya_ac(yol):
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", yol])
        elif os.name == "nt":
            os.startfile(yol)                               # noqa: yalnız Windows
        else:
            subprocess.run(["xdg-open", yol])
    except Exception:
        pass


# --------------------------------------------------------------------------
# Stil ve Görsel Arayüz Yapılandırması
# --------------------------------------------------------------------------
def ttk_stilleri_ayarla(style, renk, font_aile):
    style.theme_use("clam")

    # Genel temel yapılandırma
    style.configure(".", background=renk["bg_pencere"], foreground=renk["yazi"], font=(font_aile, 10))
    style.configure("TFrame", background=renk["bg_pencere"])
    style.configure("Card.TFrame", background=renk["bg_kart"], relief="flat")
    style.configure("CardAlt.TFrame", background=renk["bg_kart_alt"], relief="flat")

    style.configure("TLabel", background=renk["bg_pencere"], foreground=renk["yazi"], font=(font_aile, 10))
    style.configure("Card.TLabel", background=renk["bg_kart"], foreground=renk["yazi"], font=(font_aile, 10))
    style.configure("CardMuted.TLabel", background=renk["bg_kart"], foreground=renk["soluk"], font=(font_aile, 10))
    style.configure("CardBold.TLabel", background=renk["bg_kart"], foreground=renk["yazi"], font=(font_aile, 11, "bold"))

    # Birincil Aksiyon Butonu (Dönüştür / Çevir)
    style.configure("Primary.TButton", background=renk["birincil"], foreground=renk["birincil_yazi"],
                    font=(font_aile, 11, "bold"), borderwidth=0, relief="flat", padding=(18, 9))
    style.map("Primary.TButton",
              background=[("pressed", renk["birincil_hover"]),
                          ("active", renk["birincil_hover"]),
                          ("disabled", renk["cizgi"])],
              foreground=[("disabled", renk["soluk_acik"])])

    # İkincil Butonlar (Aksiyonlar)
    style.configure("Secondary.TButton", background=renk["ikincil_bg"], foreground=renk["ikincil_yazi"],
                    font=(font_aile, 9, "bold"), borderwidth=1, bordercolor=renk["ikincil_cizgi"],
                    relief="flat", padding=(10, 6))
    style.map("Secondary.TButton",
              background=[("pressed", renk["ikincil_cizgi"]),
                          ("active", renk["ikincil_hover"]),
                          ("disabled", renk["ikincil_bg"])],
              foreground=[("disabled", renk["soluk"])])

    # Hayalet Butonlar (Header / Küçük kontroller)
    style.configure("Ghost.TButton", background=renk["bg_pencere"], foreground=renk["yazi_ikincil"],
                    font=(font_aile, 9), borderwidth=1, bordercolor=renk["cizgi"],
                    relief="flat", padding=(8, 4))
    style.map("Ghost.TButton",
              background=[("active", renk["bg_kart"]), ("pressed", renk["cizgi"])],
              foreground=[("active", renk["yazi"])])

    # Açılır Liste (Combobox)
    style.configure("TCombobox", fieldbackground=renk["bg_kart"], background=renk["ikincil_bg"],
                    foreground=renk["yazi"], arrowcolor=renk["yazi"],
                    bordercolor=renk["cizgi"], lightcolor=renk["cizgi"], darkcolor=renk["cizgi"],
                    padding=(6, 4))
    style.map("TCombobox",
              fieldbackground=[("readonly", renk["bg_kart"])],
              selectbackground=[("readonly", renk["secim"])],
              selectforeground=[("readonly", renk["secim_yazi"])])

    # İlerleme Çubuğu (Progressbar)
    style.configure("Horizontal.TProgressbar", troughcolor=renk["cizgi"], background=renk["birincil"],
                    bordercolor=renk["cizgi"], lightcolor=renk["birincil"], darkcolor=renk["birincil"])

    # Ağaç / Liste Görünümü (Treeview)
    style.configure("Modern.Treeview", background=renk["bg_kart"], fieldbackground=renk["bg_kart"],
                    foreground=renk["yazi"], borderwidth=0, relief="flat",
                    lightcolor=renk["bg_kart"], darkcolor=renk["bg_kart"], bordercolor=renk["bg_kart"],
                    font=(font_aile, 10), rowheight=32)
    style.configure("Modern.Treeview.Heading", background=renk["bg_kart_alt"], foreground=renk["soluk"],
                    font=(font_aile, 9, "bold"), borderwidth=0, relief="flat",
                    lightcolor=renk["bg_kart_alt"], darkcolor=renk["bg_kart_alt"], bordercolor=renk["cizgi"],
                    padding=(6, 6))
    style.map("Modern.Treeview",
              background=[("selected", renk["secim"])],
              foreground=[("selected", renk["secim_yazi"])])
    style.map("Modern.Treeview.Heading",
              background=[("active", renk["bg_kart_alt"])],
              relief=[("active", "flat")])

    # Kaydırma Çubuğu (Scrollbar)
    style.configure("Vertical.TScrollbar", troughcolor=renk["bg_kart"], background=renk["cizgi"],
                    bordercolor=renk["bg_kart"], arrowcolor=renk["soluk"], relief="flat")
    style.map("Vertical.TScrollbar", background=[("active", renk["soluk"])])

    # Durum Bildirim Stilleri
    for tur in ("yesil", "sari", "kirmizi"):
        bg = renk[f"{tur}_bg"]
        yazi = renk[f"{tur}_yazi"]
        style.configure(f"Durum_{tur}.TFrame", background=bg)
        style.configure(f"DurumBaslik_{tur}.TLabel", background=bg, foreground=yazi, font=(font_aile, 11, "bold"))
        style.configure(f"DurumDetay_{tur}.TLabel", background=bg, foreground=yazi, font=(font_aile, 10))


class ModernListe(ttk.Treeview):
    """Modern ve şık çok sütunlu belge listesi (tk.Listbox ile tam uyumlu)."""
    def __init__(self, parent, font_aile, **kwargs):
        super().__init__(
            parent,
            columns=("durum", "dosya", "yon"),
            show="headings",
            selectmode="extended",
            style="Modern.Treeview",
            height=kwargs.pop("height", 4),
            **kwargs
        )
        self.heading("durum", text="DURUM", anchor="center")
        self.heading("dosya", text="BELGE ADI", anchor="w")
        self.heading("yon", text="DÖNÜŞÜM", anchor="center")
        self.column("durum", width=110, minwidth=95, stretch=False, anchor="center")
        self.column("dosya", width=340, minwidth=220, stretch=True, anchor="w")
        self.column("yon", width=120, minwidth=100, stretch=False, anchor="center")
        self._items = []

    def curselection(self):
        sel = set(self.selection())
        return tuple(sorted(i for i, item_id in enumerate(self._items) if item_id in sel))

    def delete(self, first, last=None):
        """tk.Listbox.delete ile birebir uyumlu: tek indeks veya aralık silme."""
        n = len(self._items)
        if n == 0:
            return

        def _parse_idx(idx):
            if idx is None:
                return None
            s = str(idx).lower()
            if s in ("end", str(tk.END).lower()):
                return n - 1
            try:
                return int(idx)
            except (ValueError, TypeError):
                return 0

        first_idx = _parse_idx(first)
        last_idx = _parse_idx(last)

        if last is None:
            if first_idx is not None and 0 <= first_idx < len(self._items):
                item_id = self._items.pop(first_idx)
                super().delete(item_id)
        else:
            if first_idx is None:
                first_idx = 0
            if last_idx is None:
                last_idx = n - 1
            if first_idx <= 0 and last_idx >= n - 1:
                for item in self.get_children():
                    super().delete(item)
                self._items.clear()
            else:
                start = max(0, min(first_idx, n - 1))
                end = max(0, min(last_idx, n - 1))
                if start <= end:
                    to_delete = self._items[start:end + 1]
                    del self._items[start:end + 1]
                    for item_id in to_delete:
                        super().delete(item_id)

    def ekle_oge(self, durum_metni, dosya_adi, yon_metni, tag=None):
        iid = super().insert("", "end", values=(durum_metni, dosya_adi, yon_metni), tags=(tag,) if tag else ())
        self._items.append(iid)
        return iid

    def guncelle_oge(self, index, durum_metni=None, tag=None):
        try:
            idx = int(index)
        except (ValueError, TypeError):
            if isinstance(index, str) and index in self._items:
                iid = index
            else:
                return
        else:
            if 0 <= idx < len(self._items):
                iid = self._items[idx]
            else:
                return

        cur_vals = list(self.item(iid, "values"))
        if durum_metni is not None:
            cur_vals[0] = durum_metni
        kwargs = {"values": cur_vals}
        if tag is not None:
            kwargs["tags"] = (tag,)
        self.item(iid, **kwargs)

    def insert(self, index, text):
        iid = super().insert("", "end", values=("", text, ""))
        self._items.append(iid)
        return iid

    def itemconfigure(self, index, **kwargs):
        pass

    def size(self):
        return len(self._items)

    def get(self, first, last=None):
        n = len(self._items)
        if n == 0:
            return "" if last is None else ()

        def _val_at(idx):
            if 0 <= idx < n:
                vals = self.item(self._items[idx], "values")
                if len(vals) > 1 and vals[1]:
                    return vals[1]
                return " ".join(str(v) for v in vals if v).strip()
            return ""

        if last is None:
            if str(first).lower() in ("end", str(tk.END).lower()):
                return _val_at(n - 1)
            try:
                return _val_at(int(first))
            except (ValueError, TypeError):
                return ""
        else:
            try:
                start = n - 1 if str(first).lower() in ("end", str(tk.END).lower()) else int(first)
                end = n - 1 if str(last).lower() in ("end", str(tk.END).lower()) else int(last)
                return tuple(_val_at(i) for i in range(max(0, start), min(n, end + 1)))
            except Exception:
                return ()

    def selection_set(self, first, last=None):
        if not self._items:
            return
        n = len(self._items)

        def _parse(idx):
            if str(idx).lower() in ("end", str(tk.END).lower()):
                return n - 1
            try:
                return int(idx)
            except (ValueError, TypeError):
                return 0

        first_idx = _parse(first)
        if last is None:
            if 0 <= first_idx < n:
                super().selection_add(self._items[first_idx])
        else:
            last_idx = _parse(last)
            start = max(0, min(first_idx, n - 1))
            end = max(0, min(last_idx, n - 1))
            if start <= end:
                to_add = [self._items[i] for i in range(start, end + 1)]
                super().selection_add(to_add)

    def selection_clear(self, first=None, last=None):
        if not self._items:
            return
        if first is None:
            super().selection_set([])
            return
        n = len(self._items)

        def _parse(idx):
            if str(idx).lower() in ("end", str(tk.END).lower()):
                return n - 1
            try:
                return int(idx)
            except (ValueError, TypeError):
                return 0

        first_idx = _parse(first)
        if last is None:
            if 0 <= first_idx < n:
                super().selection_remove(self._items[first_idx])
        else:
            last_idx = _parse(last)
            start = max(0, min(first_idx, n - 1))
            end = max(0, min(last_idx, n - 1))
            if start <= end:
                to_remove = [self._items[i] for i in range(start, end + 1)]
                super().selection_remove(to_remove)

    def see(self, index):
        if str(index).lower() in ("end", str(tk.END).lower()):
            index = len(self._items) - 1
        elif isinstance(index, str) and index.isdigit():
            index = int(index)
        if isinstance(index, int):
            if 0 <= index < len(self._items):
                super().see(self._items[index])
            return
        try:
            super().see(index)
        except Exception:
            pass


# --------------------------------------------------------------------------
# Pencereler
# --------------------------------------------------------------------------
class RaporPenceresi(tk.Toplevel):
    def __init__(self, ana, baslik, metin, renk, genislik=86, yukseklik=26):
        super().__init__(ana)
        self.title(f"{baslik} — {UYGULAMA_ADI}")
        self.transient(ana)
        self.bind("<Escape>", lambda e: self.destroy())
        font_aile = getattr(ana, "font_aile", sistem_fontu())
        mono_font = sistem_mono_fontu()
        self.configure(bg=renk["bg_pencere"])
        pencerelere_koyu_baslik_uygula(self, getattr(ana, "tema_koyu", False))

        dis = ttk.Frame(self, padding=16)
        dis.pack(fill="both", expand=True)

        ust = ttk.Frame(dis)
        ust.pack(fill="x", pady=(0, 10))
        tk.Label(ust, text="📋 Çeviri ve Doğrulama Raporu",
                 font=(font_aile, 13, "bold"),
                 bg=renk["bg_pencere"], fg=renk["yazi"]).pack(anchor="w")
        tk.Label(ust, text="Dönüştürülen belgelerin yapısal doğrulaması ve teknik detayları.",
                 font=(font_aile, 10),
                 bg=renk["bg_pencere"], fg=renk["soluk"]).pack(anchor="w", pady=(2, 0))

        kart = tk.Frame(dis, bg=renk["bg_kart"], highlightthickness=1, highlightbackground=renk["cizgi"])
        kart.pack(fill="both", expand=True)

        kutu = tk.Text(kart, width=genislik, height=yukseklik, wrap="word",
                       font=(mono_font, 10),
                       background=renk["bg_kart"], foreground=renk["yazi"],
                       selectbackground=renk["secim"], selectforeground=renk["secim_yazi"],
                       relief="flat", borderwidth=0, padx=12, pady=10)
        kaydir = ttk.Scrollbar(kart, command=kutu.yview, style="Vertical.TScrollbar")
        kutu.configure(yscrollcommand=kaydir.set)
        kaydir.pack(side="right", fill="y")
        kutu.pack(side="left", fill="both", expand=True)
        kutu.insert("1.0", metin)
        kutu.configure(state="disabled")

        alt = ttk.Frame(dis)
        alt.pack(fill="x", pady=(12, 0))
        self.btn = ttk.Button(alt, text="📋 Panoya kopyala", style="Secondary.TButton",
                              command=lambda: self.kopyala(metin))
        self.btn.pack(side="left")
        ttk.Button(alt, text="Kapat", style="Secondary.TButton", command=self.destroy).pack(side="right")
        pencereyi_ortala(self, ana)

    def kopyala(self, metin):
        self.clipboard_clear()
        self.clipboard_append(metin)
        self.btn.configure(text="✓ Kopyalandı")
        self.after(1500, lambda: self.btn.configure(text="📋 Panoya kopyala"))


class Uygulama(TEMEL_PENCERE):
    def __init__(self):
        super().__init__()
        self.title(UYGULAMA_ADI)
        ekran_boy = self.winfo_screenheight()
        ekran_en = self.winfo_screenwidth()
        baslangic_en = min(680, max(580, ekran_en - 80))
        baslangic_boy = min(720, max(580, ekran_boy - 80))
        self.geometry(f"{baslangic_en}x{baslangic_boy}")
        self.minsize(560, 420)
        self.font_aile = sistem_fontu()
        self.tema_koyu = sistem_koyu_mu()
        self.renk = PALETLER["koyu" if self.tema_koyu else "acik"]

        self._uygula_combobox_temasi()

        # İkon yükleme
        self._logo_img = None
        if sys.platform == "win32" and os.path.exists(ICO_YOLU):
            try:
                self.iconbitmap(ICO_YOLU)
            except Exception:
                pass
        try:
            if os.path.exists(PNG_YOLU):
                from PIL import Image, ImageTk
                img = Image.open(PNG_YOLU).resize((32, 32), Image.Resampling.LANCZOS)
                self._ikon_foto = ImageTk.PhotoImage(img)
                self.iconphoto(True, self._ikon_foto)
                img_logo = Image.open(PNG_YOLU).resize((34, 34), Image.Resampling.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img_logo)
        except Exception:
            pass

        self.style = ttk.Style()
        ttk_stilleri_ayarla(self.style, self.renk, self.font_aile)

        self.yollar = []
        self.sonuclar = []
        self._son_durum = None
        self._dropzone_kompakt = False
        self.calisiyor = False
        self.hedef_klasor = None                    # None = Word dosyasının yanına
        self.editor_var = bool(udf_onizle.editor_home())
        self._kur()
        try:                                        # Finder'dan uygulamaya sürükleme
            self.createcommand("::tk::mac::OpenDocument", lambda *y: self.ekle(list(y)))
        except tk.TclError:
            pass

    def _uygula_combobox_temasi(self):
        self.option_add("*TCombobox*Listbox.background", self.renk["bg_kart"])
        self.option_add("*TCombobox*Listbox.foreground", self.renk["yazi"])
        self.option_add("*TCombobox*Listbox.selectBackground", self.renk["secim"])
        self.option_add("*TCombobox*Listbox.selectForeground", self.renk["secim_yazi"])

    # ---------------------------------------------------------------- yerleşim
    def _kur(self):
        self.configure(bg=self.renk["bg_pencere"])
        pencerelere_koyu_baslik_uygula(self, self.tema_koyu)

        # ==========================================
        # 1. ÜST BAŞLIK & MARKA BARI (Tepede Sabit)
        # ==========================================
        self.ust_bar = ttk.Frame(self, padding=(18, 12, 18, 6))
        self.ust_bar.pack(side="top", fill="x")

        # ==========================================
        # 2. ALT BİLGİ (DİPTE SABİT FOOTER)
        # ==========================================
        self.footer = tk.Frame(self, bg=self.renk["bg_pencere"], padx=18, pady=8)
        self.footer.pack(side="bottom", fill="x")

        self.lbl_telif = tk.Label(
            self.footer,
            text=f"{YAZAR} · s{SURUM}",
            font=(self.font_aile, 9),
            bg=self.renk["bg_pencere"],
            fg=self.renk["soluk"]
        )
        self.lbl_telif.pack(side="left")

        self.lbl_guvenlik = tk.Label(
            self.footer,
            text="🔒 Çevrimdışı ve Yerel",
            font=(self.font_aile, 9),
            bg=self.renk["bg_pencere"],
            fg=self.renk["soluk"]
        )
        self.lbl_guvenlik.pack(side="right")

        # ==========================================
        # 3. KAYDIRILABİLİR GÖVDE (CANVAS + SCROLLBAR)
        # ==========================================
        self.govde_tasiyici = tk.Frame(self, bg=self.renk["bg_pencere"])
        self.govde_tasiyici.pack(side="top", fill="both", expand=True)

        self.canvas = tk.Canvas(
            self.govde_tasiyici,
            bg=self.renk["bg_pencere"],
            highlightthickness=0,
            bd=0
        )
        self.scrollbar = ttk.Scrollbar(
            self.govde_tasiyici,
            orient="vertical",
            command=self.canvas.yview,
            style="Vertical.TScrollbar"
        )
        self.canvas.configure(yscrollcommand=self._kaydirma_guncelle)

        self.ana_tasiyici = tk.Frame(self.canvas, bg=self.renk["bg_pencere"], padx=18, pady=6)
        self._canvas_pencere_id = self.canvas.create_window((0, 0), window=self.ana_tasiyici, anchor="nw")

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.canvas.bind("<Configure>", self._canvas_yapilandir)
        self.ana_tasiyici.bind("<Configure>", self._icerik_yapilandir)

        # Fare tekerleğiyle kaydırma desteği
        self.bind_all("<MouseWheel>", self._fare_tekeri, add="+")
        self.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"), add="+")
        self.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"), add="+")

        dis = self.ana_tasiyici

        # Sol: Logo mark + Başlık + Versiyon rozeti + Alt başlık
        self.marka_sol = ttk.Frame(self.ust_bar)
        self.marka_sol.pack(side="left", fill="y")

        self.logo_canvas = tk.Canvas(self.marka_sol, width=38, height=38,
                                     bg=self.renk["bg_pencere"], highlightthickness=0)
        self.logo_canvas.pack(side="left", padx=(0, 10))
        self._ciz_logo_mark()

        self.baslik_kutusu = ttk.Frame(self.marka_sol)
        self.baslik_kutusu.pack(side="left", fill="y")

        self.baslik_ust = ttk.Frame(self.baslik_kutusu)
        self.baslik_ust.pack(anchor="w")

        self.lbl_baslik = tk.Label(self.baslik_ust, text=UYGULAMA_ADI, font=(self.font_aile, 15, "bold"),
                                   bg=self.renk["bg_pencere"], fg=self.renk["yazi"])
        self.lbl_baslik.pack(side="left")

        self.lbl_surum_rozet = tk.Label(self.baslik_ust, text=f"v{SURUM}", font=(self.font_aile, 8, "bold"),
                                        bg=self.renk["bg_kart_alt"], fg=self.renk["soluk"],
                                        padx=6, pady=1, relief="flat")
        self.lbl_surum_rozet.pack(side="left", padx=(8, 0))

        self.lbl_alt_baslik = tk.Label(self.baslik_kutusu,
                                       text="Word ↔ UDF • Markdown ↔ UDF Güvenli ve Biçim Koruyan Çevirici",
                                       font=(self.font_aile, 10),
                                       bg=self.renk["bg_pencere"], fg=self.renk["soluk"])
        self.lbl_alt_baslik.pack(anchor="w", pady=(1, 0))

        # Sağ: Tema Butonu & Hakkında Butonu
        self.aksiyon_sag = ttk.Frame(self.ust_bar)
        self.aksiyon_sag.pack(side="right", fill="y")

        self.btn_tema = ttk.Button(self.aksiyon_sag,
                                   text="☀️ Açık" if self.tema_koyu else "🌙 Koyu",
                                   style="Ghost.TButton",
                                   command=self.tema_degistir)
        self.btn_tema.pack(side="left", padx=(0, 6))

        self.btn_hakkinda = ttk.Button(self.aksiyon_sag, text="Hakkında", style="Ghost.TButton", command=self.hakkinda_ac)
        self.btn_hakkinda.pack(side="left")

        # ==========================================
        # 2. SÜRÜKLE-BIRAK (DROPZONE) KARTI
        # ==========================================
        self.birak = self.birak_kart = tk.Frame(
            dis,
            bg=self.renk["drop_bg"],
            highlightthickness=2,
            highlightbackground=self.renk["drop_cizgi"],
            cursor="hand2"
        )
        self.birak.pack(fill="both", expand=True, pady=(0, 10))

        self.birak_ic = tk.Frame(self.birak_kart, bg=self.renk["drop_bg"], padx=20, pady=24, cursor="hand2")
        self.birak_ic.pack(fill="both", expand=True)

        self.drop_canvas = tk.Canvas(self.birak_ic, width=64, height=64,
                                     bg=self.renk["drop_bg"], highlightthickness=0, cursor="hand2")
        self.drop_canvas.pack(pady=(0, 6))
        self._ciz_drop_ikon()

        surukleme = False
        if SURUKLENEBILIR:
            try:
                for w in (self, self.canvas, self.ana_tasiyici, self.birak, self.birak_ic):
                    try:
                        w.drop_target_register(DND_FILES)
                        w.dnd_bind("<<Drop>>", self._birakildi)
                        w.dnd_bind("<<DragEnter>>", self._surukleme_girdi)
                        w.dnd_bind("<<DragLeave>>", self._surukleme_cikti)
                    except Exception:
                        pass
                surukleme = True
            except Exception:
                surukleme = False

        self.birak_yazi = tk.Label(
            self.birak_ic,
            font=(self.font_aile, 13, "bold"),
            bg=self.renk["drop_bg"],
            fg=self.renk["yazi"],
            cursor="hand2",
            text="Word, Markdown veya UDF belgelerini buraya bırakın" if surukleme
                 else "Çevrilecek Word, Markdown veya UDF belgelerini seçin"
        )
        self.birak_yazi.pack(pady=(6, 2))

        self.birak_alt_yazi = tk.Label(
            self.birak_ic,
            font=(self.font_aile, 10),
            bg=self.renk["drop_bg"],
            fg=self.renk["soluk"],
            cursor="hand2",
            text="veya bilgisayarınızdan seçmek için bu alana tıklayın"
        )
        self.birak_alt_yazi.pack(pady=(2, 0))

        # Format hap rozetleri
        self.rozet_kutusu = tk.Frame(self.birak_ic, bg=self.renk["drop_bg"], cursor="hand2")
        self.rozet_kutusu.pack(pady=(14, 0))
        self.format_rozetleri = []
        for rozet_metni in ("📄 Word (.docx)", "⚖️ UYAP (.udf)", "📝 Markdown (.md)", "📑 Pages / RTF"):
            lbl_r = tk.Label(
                self.rozet_kutusu, text=rozet_metni,
                font=(self.font_aile, 9),
                bg=self.renk["bg_kart_alt"],
                fg=self.renk["soluk"],
                padx=8, pady=3, relief="flat", cursor="hand2"
            )
            lbl_r.pack(side="left", padx=4)
            self.format_rozetleri.append(lbl_r)

        # Hover & Tıklama Olayları
        self.birak_ogeler = [self.birak, self.birak_ic, self.drop_canvas, self.birak_yazi,
                             self.birak_alt_yazi, self.rozet_kutusu] + self.format_rozetleri
        for w in self.birak_ogeler:
            w.bind("<Button-1>", lambda e: self.dosya_ekle())
            w.bind("<Enter>", self._drop_hover_gir)
            w.bind("<Leave>", self._drop_hover_cik)

        # ==========================================
        # 3. SEÇİLEN BELGELER LİSTESİ (KART)
        # ==========================================
        self.liste_cerceve = self.liste_kart = tk.Frame(
            dis,
            bg=self.renk["bg_kart"],
            highlightthickness=1,
            highlightbackground=self.renk["cizgi"],
            padx=12, pady=10
        )

        self.liste_ust = tk.Frame(self.liste_kart, bg=self.renk["bg_kart"])
        self.liste_ust.pack(fill="x", pady=(0, 8))

        self.liste_baslik_sol = tk.Frame(self.liste_ust, bg=self.renk["bg_kart"])
        self.liste_baslik_sol.pack(side="left")

        self.lbl_liste_baslik = tk.Label(
            self.liste_baslik_sol,
            text="📁 Seçilen Belgeler",
            font=(self.font_aile, 11, "bold"),
            bg=self.renk["bg_kart"],
            fg=self.renk["yazi"]
        )
        self.lbl_liste_baslik.pack(side="left")

        self.sayi_yazi = tk.Label(
            self.liste_baslik_sol,
            text="",
            font=(self.font_aile, 9, "bold"),
            bg=self.renk["bg_kart_alt"],
            fg=self.renk["soluk"],
            padx=6, pady=1
        )
        self.sayi_yazi.pack(side="left", padx=(8, 0))

        # Sağ: Kaldır ve Temizle butonları
        self.liste_butonlar = tk.Frame(self.liste_ust, bg=self.renk["bg_kart"])
        self.liste_butonlar.pack(side="right")

        self.btn_sil = ttk.Button(self.liste_butonlar, text="🗑️ Kaldır", style="Secondary.TButton", command=self.sil)
        self.btn_sil.pack(side="right")

        self.btn_temizle = ttk.Button(self.liste_butonlar, text="✕ Temizle", style="Secondary.TButton", command=self.temizle)
        self.btn_temizle.pack(side="right", padx=(0, 6))

        # Treeview ve Scrollbar taşıyıcısı
        self.tablo_kutu = tk.Frame(self.liste_kart, bg=self.renk["bg_kart"])
        self.tablo_kutu.pack(fill="both", expand=True)

        self.liste = ModernListe(self.tablo_kutu, font_aile=self.font_aile)
        self.liste_kaydir = ttk.Scrollbar(self.tablo_kutu, orient="vertical", command=self.liste.yview, style="Vertical.TScrollbar")
        self.liste.configure(yscrollcommand=self.liste_kaydir.set)

        self.liste.pack(side="left", fill="both", expand=True)
        self.liste_kaydir.pack(side="right", fill="y")
        self.liste.bind("<Double-1>", self._liste_cift_tikla)
        self.liste.bind("<Delete>", lambda e: self.sil())
        self.liste.bind("<BackSpace>", lambda e: self.sil())
        self.liste.bind("<Control-a>", self._tumunu_sec)
        self.liste.bind("<Command-a>", self._tumunu_sec)
        self.liste.bind("<Return>", lambda e: self.belge_ac())
        self.liste.bind("<Button-3>", self._sag_tik_menusu)
        self.liste.bind("<Button-2>", self._sag_tik_menusu)
        if SURUKLENEBILIR:
            try:
                self.liste.drop_target_register(DND_FILES)
                self.liste.dnd_bind("<<Drop>>", self._birakildi)
                self.liste.dnd_bind("<<DragEnter>>", self._surukleme_girdi)
                self.liste.dnd_bind("<<DragLeave>>", self._surukleme_cikti)
            except Exception:
                pass

        # ==========================================
        # 4. AYARLAR KARTI
        # ==========================================
        self.ayar = self.ayar_kart = tk.Frame(
            dis,
            bg=self.renk["bg_kart"],
            highlightthickness=1,
            highlightbackground=self.renk["cizgi"],
            padx=14, pady=10
        )

        self.lbl_ayar_baslik = tk.Label(
            self.ayar_kart,
            text="⚙️ Dönüşüm Ayarları",
            font=(self.font_aile, 11, "bold"),
            bg=self.renk["bg_kart"],
            fg=self.renk["yazi"]
        )
        self.lbl_ayar_baslik.pack(anchor="w", pady=(0, 8))

        # Word seçenekleri (sayfa no & dolgu)
        self.ayar_word = tk.Frame(self.ayar_kart, bg=self.renk["bg_kart"])
        self.ayar_word.pack(fill="x", pady=(0, 6))

        self.lbl_sayfa_no = tk.Label(self.ayar_word, text="Sayfa No:", font=(self.font_aile, 10),
                                     bg=self.renk["bg_kart"], fg=self.renk["yazi_ikincil"])
        self.lbl_sayfa_no.pack(side="left")

        self.sayfa_no = tk.StringVar(value="Word'deki gibi")
        self.cmb_sayfa = ttk.Combobox(self.ayar_word, textvariable=self.sayfa_no,
                                      values=list(SAYFA_NO), state="readonly", width=14)
        self.cmb_sayfa.pack(side="left", padx=(6, 16))

        self.lbl_dolgu = tk.Label(self.ayar_word, text="Tablo Dolgusu:", font=(self.font_aile, 10),
                                  bg=self.renk["bg_kart"], fg=self.renk["yazi_ikincil"])
        self.lbl_dolgu.pack(side="left")

        self.dolgu = tk.StringVar(value="Renk bandı")
        self.cmb_dolgu = ttk.Combobox(self.ayar_word, textvariable=self.dolgu,
                                      values=list(DOLGU), state="readonly", width=12)
        self.cmb_dolgu.pack(side="left", padx=(6, 0))

        # Kayıt yeri
        self.ayar_kayit = tk.Frame(self.ayar_kart, bg=self.renk["bg_kart"])
        self.ayar_kayit.pack(fill="x")

        self.lbl_kayit_baslik = tk.Label(self.ayar_kayit, text="Kayıt Yeri:", font=(self.font_aile, 10),
                                         bg=self.renk["bg_kart"], fg=self.renk["yazi_ikincil"])
        self.lbl_kayit_baslik.pack(side="left")

        self.hedef_yazi = tk.Label(
            self.ayar_kayit,
            text="📁 Belgenin yanına",
            font=(self.font_aile, 9),
            bg=self.renk["bg_kart_alt"],
            fg=self.renk["soluk"],
            padx=8, pady=2
        )
        self.hedef_yazi.pack(side="left", padx=(6, 8))

        self.btn_hedef_sifirla = ttk.Button(self.ayar_kayit, text="✕ Sıfırla", style="Ghost.TButton",
                                            command=self.hedef_sifirla)
        self.btn_hedef = ttk.Button(self.ayar_kayit, text="Değiştir…", style="Secondary.TButton", command=self.hedef_sec)
        self.btn_hedef.pack(side="right")

        # ==========================================
        # 5. DÖNÜŞTÜR BUTONU & İLERLEME ÇUBUĞU
        # ==========================================
        self.btn_cevir = ttk.Button(
            dis,
            text="🔄 Belgeleri Dönüştür",
            style="Primary.TButton",
            command=self.cevir_baslat
        )
        self.ilerleme = ttk.Progressbar(dis, style="Horizontal.TProgressbar", mode="determinate")

        # ==========================================
        # 6. DURUM & SONUÇ BİLDİRİM KARTI
        # ==========================================
        self.durum_kutu = tk.Frame(
            dis,
            bg=self.renk["yesil_bg"],
            highlightthickness=2,
            highlightbackground=self.renk["yesil_cizgi"],
            padx=14, pady=12
        )
        self.durum_baslik = tk.Label(
            self.durum_kutu,
            text="",
            font=(self.font_aile, 12, "bold"),
            bg=self.renk["yesil_bg"],
            fg=self.renk["yesil_yazi"]
        )
        self.durum_baslik.pack(anchor="w")

        self.durum_detay = tk.Label(
            self.durum_kutu,
            text="",
            font=(self.font_aile, 10),
            bg=self.renk["yesil_bg"],
            fg=self.renk["yesil_yazi"],
            wraplength=550,
            justify="left"
        )
        self.durum_detay.pack(anchor="w", pady=(3, 0))
        self.durum_kutu.bind("<Configure>", self._ayarla_durum_wraplength)

        # Sonuç Aksiyon Butonları
        self.sonuc_cerceve = ttk.Frame(dis)
        self.btn_rapor = ttk.Button(self.sonuc_cerceve, text="📋 Rapor", style="Secondary.TButton", command=self.rapor_ac)
        self.btn_rapor.pack(side="left")

        self.btn_klasor = ttk.Button(self.sonuc_cerceve, text="📂 Klasörde Göster", style="Secondary.TButton",
                                     command=self.klasor_ac)
        self.btn_klasor.pack(side="left", padx=8)

        self.btn_ac = ttk.Button(self.sonuc_cerceve, text="📄 Belgeyi Aç", style="Secondary.TButton",
                                 command=self.belge_ac)
        self.btn_ac.pack(side="left", padx=(0, 8))

        self.btn_onizle = ttk.Button(self.sonuc_cerceve, text="👁️ Editör Önizlemesi", style="Secondary.TButton",
                                     command=self.onizle_baslat)


        # Klavye kısayolları
        self.bind("<Control-Return>", lambda e: self.cevir_baslat())
        self.bind("<Command-Return>", lambda e: self.cevir_baslat())
        self.bind("<Control-o>", lambda e: self.dosya_ekle())
        self.bind("<Control-O>", lambda e: self.dosya_ekle())
        self.bind("<Command-o>", lambda e: self.dosya_ekle())
        self.bind("<Command-O>", lambda e: self.dosya_ekle())
        self.bind("<Control-d>", lambda e: self.tema_degistir())
        self.bind("<Command-d>", lambda e: self.tema_degistir())
        self.bind("<F1>", lambda e: self.hakkinda_ac())

        self.listeyi_ciz()

    def _ciz_logo_mark(self):
        c = self.logo_canvas
        c.delete("all")
        c.configure(bg=self.renk["bg_pencere"])
        if hasattr(self, "_logo_img") and self._logo_img:
            c.create_image(19, 19, image=self._logo_img)
        else:
            c.create_oval(2, 2, 36, 36, fill=self.renk["birincil"], outline="")
            c.create_text(19, 19, text="UDF", fill="#ffffff", font=(self.font_aile, 9, "bold"))

    def _ciz_drop_ikon(self, hover=False):
        c = self.drop_canvas
        c.delete("all")
        bg = self.renk["drop_hover"] if hover else self.renk["drop_bg"]
        c.configure(bg=bg)
        cx, cy = 32, 32
        r = 24
        circle_bg = self.renk["secim"] if not hover else (
            "#dbeafe" if self.renk["ad"] == "Açık" else "#1e3a8a"
        )
        circle_border = self.renk["cizgi_odak"] if hover else self.renk["cizgi"]
        c.create_oval(cx - r, cy - r, cx + r, cy + r, fill=circle_bg, outline=circle_border, width=2)
        ok_renk = self.renk["birincil"] if not hover else self.renk["cizgi_odak"]
        c.create_line(cx, cy + 9, cx, cy - 8, fill=ok_renk, width=2, capstyle="round")
        c.create_polygon(cx - 7, cy - 3, cx, cy - 12, cx + 7, cy - 3, fill=ok_renk, outline="")
        c.create_line(cx - 10, cy + 13, cx + 10, cy + 13, fill=ok_renk, width=2, capstyle="round")

    def _ayarla_dropzone_modu(self, kompakt=False):
        self._dropzone_kompakt = kompakt
        if kompakt:
            self.rozet_kutusu.pack_forget()
            self.birak_alt_yazi.pack_forget()
            self.drop_canvas.pack_forget()
            self.birak.pack_configure(fill="x", expand=False)
            self.birak_ic.pack_configure(fill="x", expand=False)
            self.birak_ic.configure(pady=8, padx=14)
            self.birak_yazi.configure(
                text="➕ Daha fazla belge eklemek için buraya sürükleyin veya tıklayın",
                font=(self.font_aile, 10, "bold")
            )
        else:
            self.birak.pack_configure(fill="both", expand=True)
            self.birak_ic.pack_configure(fill="both", expand=True)
            self.birak_ic.configure(pady=24, padx=20)
            self.drop_canvas.pack(pady=(0, 6))
            surukleme = SURUKLENEBILIR
            self.birak_yazi.configure(
                font=(self.font_aile, 13, "bold"),
                text="Word, Markdown veya UDF belgelerini buraya bırakın" if surukleme
                     else "Çevrilecek Word, Markdown veya UDF belgelerini seçin"
            )
            self.birak_alt_yazi.pack(pady=(4, 0))
            self.rozet_kutusu.pack(pady=(16, 0))
            self._ciz_drop_ikon(hover=False)

    def _ayarla_durum_wraplength(self, event=None):
        try:
            if hasattr(self, "durum_detay") and hasattr(self, "durum_kutu"):
                w = self.durum_kutu.winfo_width()
                if w > 100:
                    self.durum_detay.configure(wraplength=max(320, w - 40))
        except Exception:
            pass

    def _canvas_yapilandir(self, event):
        if hasattr(self, "canvas") and hasattr(self, "_canvas_pencere_id"):
            self.canvas.itemconfig(self._canvas_pencere_id, width=event.width)

    def _icerik_yapilandir(self, event=None):
        if hasattr(self, "canvas"):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _kaydirma_guncelle(self, first, last):
        if hasattr(self, "scrollbar"):
            self.scrollbar.set(first, last)
            try:
                f, l = float(first), float(last)
                if f <= 0.0 and l >= 1.0:
                    if bool(self.scrollbar.winfo_manager()):
                        self.scrollbar.pack_forget()
                else:
                    if not bool(self.scrollbar.winfo_manager()):
                        self.scrollbar.pack(side="right", fill="y", before=self.canvas)
            except Exception:
                pass

    def _fare_tekeri(self, event):
        try:
            if not hasattr(self, "canvas"):
                return
            x, y = event.x_root, event.y_root
            w = self.winfo_containing(x, y)
            liste_uzerinde = False
            cur = w
            while cur:
                if hasattr(self, "liste") and cur == self.liste:
                    liste_uzerinde = True
                    break
                cur = getattr(cur, "master", None)

            if not liste_uzerinde or (hasattr(self, "liste") and self.liste.size() <= 4):
                if sys.platform == "darwin":
                    delta = -1 * int(event.delta)
                else:
                    delta = int(-1 * (event.delta / 120))
                self.canvas.yview_scroll(delta, "units")
                return "break"
        except Exception:
            pass

    def _gorunur_yap(self, widget):
        """İlgili widget'ın görünür olmasını sağlamak için canvas'ı gerektiğinde kaydırır."""
        try:
            if not hasattr(self, "canvas") or not bool(widget.winfo_manager()):
                return
            self.update_idletasks()
            bbox = self.canvas.bbox("all")
            if not bbox:
                return
            top_y, bot_y = bbox[1], bbox[3]
            total_h = bot_y - top_y
            c_h = self.canvas.winfo_height()
            if total_h <= c_h or c_h <= 0:
                return

            w_y = widget.winfo_y()
            w_h = widget.winfo_height()
            w_bottom = w_y + w_h

            current_top_ratio, current_bot_ratio = self.canvas.yview()
            visible_top = current_top_ratio * total_h
            visible_bot = current_bot_ratio * total_h

            if w_bottom > visible_bot:
                target = (w_bottom - c_h + 16) / total_h
                self.canvas.yview_moveto(min(1.0, max(0.0, target)))
            elif w_y < visible_top:
                target = max(0.0, (w_y - 16) / total_h)
                self.canvas.yview_moveto(target)
        except Exception:
            pass

    def _tumunu_sec(self, event=None):
        self.liste.selection_set(self.liste.get_children())
        return "break"

    def _sag_tik_menusu(self, event):
        row_id = self.liste.identify_row(event.y)
        if row_id:
            if row_id not in self.liste.selection():
                self.liste.selection_set(row_id)
        if not self.liste.curselection():
            return

        menu = tk.Menu(self, tearoff=0, bg=self.renk["bg_kart"], fg=self.renk["yazi"],
                       activebackground=self.renk["secim"], activeforeground=self.renk["secim_yazi"],
                       relief="flat", bd=1)
        r = self._secili_sonuc()
        if r and r["tamam"]:
            menu.add_command(label="📄 Çevrilen Belgeyi Aç", command=self.belge_ac)
            menu.add_command(label="📂 Çıktı Klasöründe Göster", command=self.klasor_ac)
            if self.editor_var and r["yon"] == "word>udf":
                menu.add_command(label="👁️ Editör Önizlemesi", command=self.onizle_baslat)
            menu.add_separator()
        else:
            sec = self.liste.curselection()
            if sec:
                menu.add_command(label="📄 Kaynak Belgeyi Aç", command=lambda: dosya_ac(self.yollar[sec[0]]))
                menu.add_command(label="📂 Klasörde Göster", command=lambda: klasorde_goster(self.yollar[sec[0]]))
                menu.add_separator()

        menu.add_command(label="🗑️ Listeden Kaldır", command=self.sil)
        menu.add_command(label="✕ Listeyi Temizle", command=self.temizle)
        if self.sonuclar:
            menu.add_separator()
            menu.add_command(label="📋 Çeviri Raporunu Gör", command=self.rapor_ac)

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _drop_hover_gir(self, event=None):
        if self.calisiyor:
            return
        bg = self.renk["drop_hover"]
        self.birak_kart.configure(bg=bg, highlightbackground=self.renk["drop_cizgi_hover"])
        self.birak_ic.configure(bg=bg)
        self.birak_yazi.configure(bg=bg, fg=self.renk["yazi"])
        self.birak_alt_yazi.configure(bg=bg, fg=self.renk["soluk"])
        self.rozet_kutusu.configure(bg=bg)
        if not self._dropzone_kompakt:
            self._ciz_drop_ikon(hover=True)

    def _drop_hover_cik(self, event=None):
        if self.calisiyor:
            return
        bg = self.renk["drop_bg"]
        self.birak_kart.configure(bg=bg, highlightbackground=self.renk["drop_cizgi"])
        self.birak_ic.configure(bg=bg)
        self.birak_yazi.configure(bg=bg, fg=self.renk["yazi"])
        self.birak_alt_yazi.configure(bg=bg, fg=self.renk["soluk"])
        self.rozet_kutusu.configure(bg=bg)
        if not self._dropzone_kompakt:
            self._ciz_drop_ikon(hover=False)

    def _surukleme_girdi(self, event=None):
        self._drop_hover_gir()

    def _surukleme_cikti(self, event=None):
        self._drop_hover_cik()

    def tema_degistir(self):
        self.tema_koyu = not self.tema_koyu
        self.renk = PALETLER["koyu" if self.tema_koyu else "acik"]
        self._tema_guncelle()

    def _tema_guncelle(self):
        self.configure(bg=self.renk["bg_pencere"])
        pencerelere_koyu_baslik_uygula(self, self.tema_koyu)
        self._uygula_combobox_temasi()
        self.style = ttk.Style()
        ttk_stilleri_ayarla(self.style, self.renk, self.font_aile)

        self.btn_tema.configure(text="☀️ Açık" if self.tema_koyu else "🌙 Koyu")
        if hasattr(self, "govde_tasiyici"):
            self.govde_tasiyici.configure(bg=self.renk["bg_pencere"])
        if hasattr(self, "canvas"):
            self.canvas.configure(bg=self.renk["bg_pencere"])
        if hasattr(self, "ana_tasiyici"):
            self.ana_tasiyici.configure(bg=self.renk["bg_pencere"])
        if hasattr(self, "ust_bar"):
            self.ust_bar.configure(style="TFrame")
        self.lbl_baslik.configure(bg=self.renk["bg_pencere"], fg=self.renk["yazi"])
        self.lbl_surum_rozet.configure(bg=self.renk["bg_kart_alt"], fg=self.renk["soluk"])
        self.lbl_alt_baslik.configure(bg=self.renk["bg_pencere"], fg=self.renk["soluk"])

        bg_drop = self.renk["drop_bg"]
        self.birak_kart.configure(bg=bg_drop, highlightbackground=self.renk["drop_cizgi"])
        self.birak_ic.configure(bg=bg_drop)
        self.birak_yazi.configure(bg=bg_drop, fg=self.renk["yazi"])
        self.birak_alt_yazi.configure(bg=bg_drop, fg=self.renk["soluk"])
        self.rozet_kutusu.configure(bg=bg_drop)

        self._ciz_logo_mark()
        self._ciz_drop_ikon()

        self.liste_kart.configure(bg=self.renk["bg_kart"], highlightbackground=self.renk["cizgi"])
        if hasattr(self, "liste_ust"):
            self.liste_ust.configure(bg=self.renk["bg_kart"])
            self.liste_baslik_sol.configure(bg=self.renk["bg_kart"])
            self.liste_butonlar.configure(bg=self.renk["bg_kart"])
            self.tablo_kutu.configure(bg=self.renk["bg_kart"])
        self.lbl_liste_baslik.configure(bg=self.renk["bg_kart"], fg=self.renk["yazi"])
        self.sayi_yazi.configure(bg=self.renk["bg_kart_alt"], fg=self.renk["soluk"])

        self.ayar_kart.configure(bg=self.renk["bg_kart"], highlightbackground=self.renk["cizgi"])
        self.lbl_ayar_baslik.configure(bg=self.renk["bg_kart"], fg=self.renk["yazi"])
        self.ayar_word.configure(bg=self.renk["bg_kart"])
        self.lbl_sayfa_no.configure(bg=self.renk["bg_kart"], fg=self.renk["yazi_ikincil"])
        self.lbl_dolgu.configure(bg=self.renk["bg_kart"], fg=self.renk["yazi_ikincil"])
        self.ayar_kayit.configure(bg=self.renk["bg_kart"])
        self.lbl_kayit_baslik.configure(bg=self.renk["bg_kart"], fg=self.renk["yazi_ikincil"])
        self.hedef_yazi.configure(bg=self.renk["bg_kart_alt"], fg=self.renk["soluk"])

        self.footer.configure(bg=self.renk["bg_pencere"])
        self.lbl_telif.configure(bg=self.renk["bg_pencere"], fg=self.renk["soluk"])
        self.lbl_guvenlik.configure(bg=self.renk["bg_pencere"], fg=self.renk["soluk"])

        for lbl_r in self.format_rozetleri:
            lbl_r.configure(bg=self.renk["bg_kart_alt"], fg=self.renk["soluk"])

        if self._son_durum and bool(self.durum_kutu.winfo_manager()):
            tur, baslik, detay = self._son_durum
            renk_bg = self.renk.get(f"{tur}_bg", self.renk["bg_kart"])
            renk_cizgi = self.renk.get(f"{tur}_cizgi", self.renk[tur])
            renk_yazi = self.renk.get(f"{tur}_yazi", self.renk[tur])
            ikon = "✓ " if tur == "yesil" else ("⚠️ " if tur == "sari" else "✕ ")
            self.durum_kutu.configure(highlightbackground=renk_cizgi, bg=renk_bg)
            self.durum_baslik.configure(text=f"{ikon}{baslik}", bg=renk_bg, fg=renk_yazi)
            self.durum_detay.configure(text=detay, bg=renk_bg, fg=renk_yazi)

        self.listeyi_ciz()

    def _liste_cift_tikla(self, event=None):
        self.belge_ac()

    # ------------------------------------------------------------ dosya işleri
    def _birakildi(self, olay):
        self._drop_hover_cik()
        ham = getattr(olay, "data", "") or ""
        try:
            ham_yollar = self.tk.splitlist(ham)
        except Exception:
            ham_yollar = [ham]
        yollar = []
        for y in ham_yollar:
            temiz = str(y).strip()
            if temiz.startswith("{") and temiz.endswith("}"):
                temiz = temiz[1:-1].strip()
            if temiz:
                yollar.append(temiz)
        self.ekle(yollar)

    def dosya_ekle(self):
        if self.calisiyor:
            return
        secilen = filedialog.askopenfilenames(
            title="Word, Pages, Markdown ya da UDF belgelerini seçin",
            filetypes=[("Word, Pages, Markdown ve UDF belgeleri", "*.docx *.doc *.rtf *.odt *.pages *.md *.markdown *.udf"),
                       ("Word belgesi", "*.docx *.doc *.rtf *.odt"), ("Markdown belgesi", "*.md *.markdown"),
                       ("Pages belgesi", "*.pages"), ("UYAP belgesi", "*.udf"), ("Tüm dosyalar", "*.*")])
        self.ekle(list(secilen))

    def ekle(self, yeni):
        if self.calisiyor:
            return
        atlanan = []
        for y in yeni:
            y = os.path.abspath(str(y))
            y = y.rstrip("/\\")
            if os.path.isdir(y) and not y.lower().endswith(".pages"):   # eski .pages belgesi bir klasördür
                self.ekle([os.path.join(y, f) for f in sorted(os.listdir(y))
                           if f.lower().endswith(KABUL) and not f.startswith("~$")])
                continue
            if not y.lower().endswith(KABUL) or os.path.basename(y).startswith("~$"):
                atlanan.append(os.path.basename(y))
                continue
            if y not in self.yollar and os.path.exists(y):
                self.yollar.append(y)
        self.sifirla()
        self.listeyi_ciz()
        if atlanan:
            self.durum_goster("sari", "Bazı dosyalar eklenmedi",
                              "Yalnız Word (.docx, .doc, .rtf, .odt), Pages, Markdown (.md) ve UDF belgeleri çevrilir: "
                              + ", ".join(atlanan[:4]) + (" …" if len(atlanan) > 4 else ""))

    def sil(self):
        if self.calisiyor:
            return
        for i in reversed(self.liste.curselection()):
            del self.yollar[i]
        self.sifirla()
        self.listeyi_ciz()

    def temizle(self):
        if self.calisiyor:
            return
        self.yollar = []
        self.sifirla()
        self.listeyi_ciz()

    def listeyi_ciz(self):
        self.liste.delete(0, "end")
        durum = {r["kaynak"]: r for r in self.sonuclar}
        for y in self.yollar:
            r = durum.get(y)
            ext = os.path.splitext(y)[1].lower()
            if ext == ".udf":
                dosya_ikon = "⚖️"
                yon = "UDF → Word"
            elif ext in (".md", ".markdown"):
                dosya_ikon = "📝"
                yon = "MD → UDF"
            elif ext == ".pages":
                dosya_ikon = "📑"
                yon = "Pages → UDF"
            else:
                dosya_ikon = "📄"
                yon = "Word → UDF"

            ad = os.path.basename(y)
            if r is None:
                durum_ikon = "⏳ Bekliyor"
                tag = "bekliyor"
            elif r["tamam"]:
                durum_ikon = "✓ Çevrildi"
                tag = "basarili"
            else:
                durum_ikon = "✕ Hata"
                tag = "hatali"

            self.liste.ekle_oge(durum_ikon, f"{dosya_ikon}  {ad}", yon, tag=tag)

        self.liste.tag_configure("cevrimde", foreground=self.renk["birincil"])
        self.liste.tag_configure("basarili", foreground=self.renk["yesil"])
        self.liste.tag_configure("hatali", foreground=self.renk["kirmizi"])
        self.liste.tag_configure("bekliyor", foreground=self.renk["soluk"])

        if self.yollar:
            self._ayarla_dropzone_modu(kompakt=True)
            self.liste_cerceve.pack(fill="x", expand=False, after=self.birak, pady=(0, 10))
            self.ayar.pack(fill="x", after=self.liste_cerceve, pady=(0, 10))
            self.btn_cevir.pack(fill="x", pady=(0, 10), after=self.ayar)
            self.sayi_yazi.configure(text=f"{len(self.yollar)} belge")
            btn_metni = f"🔄 Yeniden Çevir ({len(self.yollar)} Belge)" if self.sonuclar else f"🔄 Çevir ({len(self.yollar)} Belge)"
            self.btn_cevir.configure(text=btn_metni)
            if any(not y.lower().endswith(".udf") for y in self.yollar):
                self.ayar_word.pack(fill="x", before=self.ayar_kayit, pady=(0, 6))
            else:
                self.ayar_word.pack_forget()
            self._gorunur_yap(self.btn_cevir)
        else:
            self._ayarla_dropzone_modu(kompakt=False)
            for w in (self.liste_cerceve, self.ayar, self.btn_cevir):
                w.pack_forget()

    def sifirla(self):
        self.sonuclar = []
        self._son_durum = None
        for w in (self.durum_kutu, self.sonuc_cerceve, self.ilerleme):
            w.pack_forget()

    def hedef_sec(self):
        k = filedialog.askdirectory(title="Çevrilen belgelerin kaydedileceği klasör")
        if k:
            self.hedef_klasor = k
            self.hedef_yazi.configure(text=k if len(k) < 38 else "…" + k[-36:])
            self.btn_hedef_sifirla.pack(side="right", padx=(0, 6), before=self.btn_hedef)

    def hedef_sifirla(self):
        self.hedef_klasor = None
        self.hedef_yazi.configure(text="📁 Belgenin yanına")
        self.btn_hedef_sifirla.pack_forget()

    def durum_goster(self, tur, baslik, detay):
        self._son_durum = (tur, baslik, detay)
        renk_bg = self.renk.get(f"{tur}_bg", self.renk["bg_kart"])
        renk_cizgi = self.renk.get(f"{tur}_cizgi", self.renk[tur])
        renk_yazi = self.renk.get(f"{tur}_yazi", self.renk[tur])
        ikon = "✓ " if tur == "yesil" else ("⚠️ " if tur == "sari" else "✕ ")

        self.durum_kutu.configure(highlightbackground=renk_cizgi, bg=renk_bg)
        self.durum_baslik.configure(text=f"{ikon}{baslik}", bg=renk_bg, fg=renk_yazi)
        self.durum_detay.configure(text=detay, bg=renk_bg, fg=renk_yazi)
        hedef_onceki = self.btn_cevir if bool(self.btn_cevir.winfo_manager()) else self.birak
        self.durum_kutu.pack(fill="x", pady=(10, 0), after=hedef_onceki)
        self._ayarla_durum_wraplength()
        self._gorunur_yap(self.durum_kutu)

    # ---------------------------------------------------------------- çevirme
    def cevir_baslat(self):
        if self.calisiyor or not self.yollar:
            return
        var_olan = []
        for y in self.yollar:
            if y.lower().endswith(".udf"):
                continue
            hedef = os.path.join(self.hedef_klasor or os.path.dirname(y),
                                 os.path.splitext(os.path.basename(y))[0] + ".udf")
            if os.path.exists(hedef):
                var_olan.append(os.path.basename(hedef))
        if var_olan and not messagebox.askyesno(
                UYGULAMA_ADI, "Şu UDF dosyaları zaten var, üzerine yazılsın mı?\n\n"
                + "\n".join(var_olan[:6]) + ("\n…" if len(var_olan) > 6 else "")):
            return
        self.sifirla()
        self.calisiyor = True
        self.btn_cevir.configure(state="disabled", text="Dönüştürülüyor… ⏳")
        self.ilerleme.configure(maximum=len(self.yollar), value=0)
        self.ilerleme.pack(fill="x", pady=(0, 10), after=self.btn_cevir)
        ayar = (self.hedef_klasor, SAYFA_NO[self.sayfa_no.get()], DOLGU[self.dolgu.get()])
        threading.Thread(target=self._cevir_isi, args=(list(self.yollar),) + ayar, daemon=True).start()

    def _cevir_isi(self, yollar, klasor, sayfa_no, dolgu):
        sonuclar = []
        for i, y in enumerate(yollar, 1):
            self.after(0, lambda idx=i-1: self._adim_basladi(idx))
            res = belge_cevir(y, klasor, sayfa_no, dolgu)
            sonuclar.append(res)
            self.after(0, lambda n=i, idx=i-1, r=res: self._adim_bitti(n, idx, r))
        self.after(0, lambda: self._cevir_bitti(sonuclar))

    def _adim_basladi(self, idx):
        if isinstance(idx, str) and idx in self.yollar:
            idx = self.yollar.index(idx)
        self.liste.guncelle_oge(idx, "⚡ Çevriliyor…", tag="cevrimde")
        self.liste.see(idx)

    def _adim_bitti(self, n, idx=None, sonuc=None):
        if sonuc is None and isinstance(n, dict):
            sonuc = n
            idx = 0
            if sonuc.get("kaynak") in self.yollar:
                idx = self.yollar.index(sonuc["kaynak"])
            n = idx + 1
        elif isinstance(idx, str) and idx in self.yollar:
            idx = self.yollar.index(idx)
        if isinstance(n, int):
            self.ilerleme.configure(value=n)
        if sonuc:
            tag = "basarili" if sonuc.get("tamam") else "hatali"
            durum_metni = "✓ Çevrildi" if sonuc.get("tamam") else "✕ Hata"
            self.liste.guncelle_oge(idx, durum_metni, tag=tag)

    def _cevir_bitti(self, sonuclar):
        self.calisiyor = False
        self.sonuclar = sonuclar
        self.btn_cevir.configure(state="normal")
        self.ilerleme.pack_forget()
        self.listeyi_ciz()
        iyi = [r for r in sonuclar if r["tamam"]]
        kotu = [r for r in sonuclar if not r["tamam"]]
        self._sonuc_yaz(iyi, kotu, 4)
        self.sonuc_cerceve.pack(fill="x", pady=(10, 0), after=self.durum_kutu)
        for w in (self.btn_rapor, self.btn_klasor, self.btn_ac, self.btn_onizle):
            w.pack_forget()
        self.btn_rapor.pack(side="left")
        if iyi:
            self.btn_klasor.pack(side="left", padx=6)
            self.btn_ac.pack(side="left", padx=(0, 6))
            if self.editor_var and any(r["yon"] == "word>udf" for r in iyi):
                self.btn_onizle.pack(side="left")
                self.btn_onizle.configure(state="normal")
        self._sigdir()
        if self.winfo_reqheight() > self.winfo_screenheight() - 90:
            self._sonuc_yaz(iyi, kotu, 2)
            self._sigdir()
        self._gorunur_yap(self.sonuc_cerceve if bool(self.sonuc_cerceve.winfo_manager()) else self.durum_kutu)

    def _sonuc_yaz(self, iyi, kotu, en_cok):
        """Sonuç kutusunun metni; uyarılar sade dille, en çok en_cok madde."""
        gosterilecek, kalan = sade_uyarilar(iyi, en_cok)
        yonler = {r["yon"] for r in iyi}
        farklar = ""
        if gosterilecek:
            bas = ("Word'den farklı olan yerler:" if yonler == {"word>udf"} else
                   "Editör'dekinden farklı olan yerler:" if yonler == {"udf>word"} else "Dikkat edilecek yerler:")
            farklar = ("\n\n" + bas + "\n" + "\n".join("•  " + u for u in gosterilecek)
                       + (f"\n•  … ve {kalan} husus daha (Rapor'da)" if kalan else ""))
        son = []
        if "word>udf" in yonler:
            son.append("UYAP'a yüklemeden önce UDF'yi Editör'de açıp bir kez kontrol edin.")
        if "udf>word" in yonler:
            yeni_ad = [os.path.basename(r["cikti"]) for r in iyi if r["yon"] == "udf>word" and "(UDF'den" in r["cikti"]]
            if yeni_ad:
                son.append("Aynı adda bir Word belgesi vardı; üzerine yazılmadı, yeni belge \""
                           + yeni_ad[0] + "\" adıyla kaydedildi.")
            son.append("Word'de satır ve sayfa sonları Editör'dekinden biraz farklı düşebilir.")
        if len(son) > 2:
            son = son[:2]
        if iyi and not kotu:
            baslik = "Belge çevrildi" if len(iyi) == 1 else f"{len(iyi)} belge çevrildi"
            detay = "Doğrulama geçti; metin eksiksiz." + farklar + ("\n\n" + "\n".join(son) if son else "")
            self.durum_goster("yesil" if not gosterilecek else "sari", baslik, detay)
        elif iyi:
            self.durum_goster("sari", f"{len(iyi)} belge çevrildi, {len(kotu)} belge çevrilemedi",
                              "İlk sorun: " + kotu[0]["hata"] + farklar + "\n\nAyrıntı Rapor'da.")
        else:
            self.durum_goster("kirmizi", "Çevrilemedi",
                              kotu[0]["hata"] if kotu else "Bilinmeyen bir sorun oluştu.")

    def _sigdir(self):
        """Sonuç kutusu uzunsa pencereyi içeriğe göre uzatır (düğmeler alta taşıp kaybolmasın)."""
        self.update_idletasks()
        en = max(self.winfo_width(), self.winfo_reqwidth())
        boy = min(max(self.winfo_height(), self.winfo_reqheight()), self.winfo_screenheight() - 90)
        if (en, boy) != (self.winfo_width(), self.winfo_height()):
            self.geometry(f"{en}x{boy}")

    # ------------------------------------------------------------------ sonuç
    def rapor_ac(self):
        if self.sonuclar:
            RaporPenceresi(self, "Çeviri raporu", rapor_metni(self.sonuclar), self.renk)

    def _secili_sonuc(self, yalniz_tamam=False):
        sec = self.liste.curselection()
        if sec and sec[0] < len(self.yollar):
            y = self.yollar[sec[0]]
            for r in self.sonuclar:
                if r["kaynak"] == y:
                    if yalniz_tamam and not r["tamam"]:
                        return None
                    return r
        iyi = [r for r in self.sonuclar if r["tamam"]]
        if yalniz_tamam:
            return iyi[0] if iyi else None
        return iyi[0] if iyi else (self.sonuclar[0] if self.sonuclar else None)

    def klasor_ac(self):
        if not self.sonuclar:
            sec = self.liste.curselection()
            if sec and sec[0] < len(self.yollar):
                klasorde_goster(self.yollar[sec[0]])
            return
        r = self._secili_sonuc()
        if r:
            if r["tamam"]:
                klasorde_goster(r["cikti"])
            else:
                klasorde_goster(r["kaynak"])

    def belge_ac(self):
        if not self.sonuclar:
            sec = self.liste.curselection()
            if sec and sec[0] < len(self.yollar):
                dosya_ac(self.yollar[sec[0]])
            return
        r = self._secili_sonuc()
        if r:
            if r["tamam"]:
                dosya_ac(r["cikti"])
            else:
                messagebox.showwarning(
                    UYGULAMA_ADI,
                    f"'{os.path.basename(r['kaynak'])}' belgesi çevrilemediği için çıktısı bulunmuyor.\n\n"
                    f"Sorun detayı: {r['hata']}"
                )

    def onizle_baslat(self):
        r = self._secili_sonuc(yalniz_tamam=True)
        if r and r["yon"] != "word>udf":
            r = next((x for x in self.sonuclar if x["tamam"] and x["yon"] == "word>udf"), None)
        if not r or self.calisiyor:
            return
        self.calisiyor = True
        self.btn_onizle.configure(state="disabled", text="Çiziliyor…")
        threading.Thread(target=self._onizle_isi, args=(r["cikti"],), daemon=True).start()

    def _onizle_isi(self, udf):
        try:
            klasor = tempfile.mkdtemp(prefix="udf_onizleme_")
            _, sayfalar = udf_onizle.onizle(udf, klasor, olcek=1.4)
            hata = None
        except Exception as e:
            sayfalar, hata = [], str(e)
        self.after(0, lambda: self._onizle_bitti(sayfalar, hata))

    def _onizle_bitti(self, sayfalar, hata):
        self.calisiyor = False
        self.btn_onizle.configure(state="normal", text="👁️ Editör Önizlemesi")
        if hata or not sayfalar:
            messagebox.showwarning(UYGULAMA_ADI, "Önizleme çizilemedi.\n\n" + (hata or ""))
            return
        if sys.platform == "darwin":
            subprocess.run(["open"] + sayfalar)
        else:
            dosya_ac(sayfalar[0])
            if len(sayfalar) > 1:
                klasorde_goster(sayfalar[0])

    # --------------------------------------------------------------- hakkında
    def hakkinda_ac(self):
        p = tk.Toplevel(self)
        p.title(f"Hakkında — {UYGULAMA_ADI}")
        p.transient(self)
        p.resizable(False, False)
        p.bind("<Escape>", lambda e: p.destroy())
        p.configure(bg=self.renk["bg_pencere"])
        pencerelere_koyu_baslik_uygula(p, self.tema_koyu)

        c = tk.Frame(p, bg=self.renk["bg_kart"], highlightthickness=1,
                     highlightbackground=self.renk["cizgi"], padx=24, pady=20)
        c.pack(fill="both", expand=True, padx=16, pady=16)

        # Logo mark
        logo = tk.Canvas(c, width=50, height=50, bg=self.renk["bg_kart"], highlightthickness=0)
        logo.pack(pady=(0, 6))
        if hasattr(self, "_logo_img") and self._logo_img:
            logo.create_image(25, 25, image=self._logo_img)
        else:
            logo.create_oval(3, 3, 47, 47, fill=self.renk["birincil"], outline="")
            logo.create_text(25, 25, text="UDF", fill="#ffffff", font=(self.font_aile, 12, "bold"))

        tk.Label(c, text=UYGULAMA_ADI, font=(self.font_aile, 16, "bold"),
                 bg=self.renk["bg_kart"], fg=self.renk["yazi"]).pack()
        tk.Label(c, text=f"Sürüm {SURUM}", font=(self.font_aile, 10, "bold"),
                 bg=self.renk["bg_kart_alt"], fg=self.renk["soluk"], padx=8, pady=2).pack(pady=(4, 12))

        tk.Label(c, justify="center", wraplength=420, font=(self.font_aile, 10),
                 bg=self.renk["bg_kart"], fg=self.renk["yazi_ikincil"], text=(
            "Word (.docx) belgelerini; tabloları, görselleri, listeleri, üstbilgi ve "
            "altbilgisiyle UYAP Doküman Editörü'nün UDF biçimine, UDF belgelerini de "
            "aynı sadakatle Word'e çevirir. Markdown (.md), .doc, .rtf, .odt ve Pages "
            "biçimleri de desteklenir.\n\n"
            "🔒 Belgeler bilgisayarınızdan asla çıkmaz; tüm işlemler yerel olarak yapılır. "
            "Program kendiliğinden internete bağlanmaz.")).pack()

        tk.Frame(c, height=1, bg=self.renk["cizgi"]).pack(fill="x", pady=14)

        tk.Label(c, text=YAZAR, font=(self.font_aile, 11, "bold"),
                 bg=self.renk["bg_kart"], fg=self.renk["yazi"]).pack()
        tk.Label(c, text=YAZAR_EK, font=(self.font_aile, 10),
                 bg=self.renk["bg_kart"], fg=self.renk["soluk"]).pack(pady=(2, 0))
        tk.Label(c, text=LISANS, font=(self.font_aile, 9),
                 bg=self.renk["bg_kart"], fg=self.renk["soluk_acik"]).pack(pady=(6, 0))

        editor_durum = ("✓ UYAP Doküman Editörü bulundu (Önizleme etkin)" if self.editor_var
                        else "UYAP Doküman Editörü bulunamadı (Önizleme pasif)")
        tk.Label(c, text=editor_durum, font=(self.font_aile, 9),
                 bg=self.renk["bg_kart_alt"], fg=self.renk["soluk"], padx=8, pady=3).pack(pady=(10, 0))

        tk.Label(c, text="Bu program UYAP veya Adalet Bakanlığı ile resmi bir bağlantıya sahip değildir.",
                 font=(self.font_aile, 8), bg=self.renk["bg_kart"], fg=self.renk["soluk_acik"]).pack(pady=(4, 0))

        if DEPO:
            guncelleme_kutusu = tk.Frame(c, bg=self.renk["bg_kart"])
            guncelleme_kutusu.pack(pady=(12, 0))
            yazi = tk.Label(guncelleme_kutusu, text="", font=(self.font_aile, 9),
                            bg=self.renk["bg_kart"], fg=self.renk["soluk"])
            btn_indir = ttk.Button(guncelleme_kutusu, text="İndirme Sayfasını Aç", style="Secondary.TButton",
                                   command=lambda: webbrowser.open(SURUM_SAYFA))
            dugme = ttk.Button(guncelleme_kutusu, text="🔄 Güncellemeleri Denetle", style="Secondary.TButton")
            dugme.configure(command=lambda: self.guncelleme_kontrol(p, dugme, yazi, btn_indir))
            dugme.pack()
            yazi.pack(pady=(4, 0))

        ttk.Button(c, text="Kapat", style="Secondary.TButton", command=p.destroy).pack(pady=(14, 0))
        pencereyi_ortala(p, self)

    def guncelleme_kontrol(self, pencere, dugme, yazi, btn_indir):
        dugme.configure(state="disabled", text="Denetleniyor…")

        def bitti(etiket, hata):
            if not pencere.winfo_exists():
                return
            dugme.configure(state="normal", text="🔄 Güncellemeleri Denetle")
            if hata:
                yazi.configure(text="Denetlenemedi (internet bağlantısı yok olabilir).")
                btn_indir.pack_forget()
            elif surum_sayilari(etiket) > surum_sayilari(SURUM):
                yazi.configure(text=f"Yeni sürüm var: {etiket}")
                btn_indir.pack(pady=(6, 0))
            else:
                yazi.configure(text="En güncel sürümü kullanıyorsunuz.")
                btn_indir.pack_forget()

        def is_parcacigi():
            try:
                e, h = son_surumu_sor(), None
            except (urllib.error.URLError, OSError, ValueError) as ex:
                e, h = "", ex
            self.after(0, lambda: bitti(e, h))
        threading.Thread(target=is_parcacigi, daemon=True).start()


# --------------------------------------------------------------------------
# Öz sınama
# --------------------------------------------------------------------------
def _ornek_docx(yol):
    """Sınama için küçük bir Word belgesi: biçimli paragraf + 2x2 tablo + PNG görsel."""
    import struct
    import zlib

    def parca(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    ham = b"".join(b"\x00" + bytes((31, 58, 95)) * 24 for _ in range(12))
    png = (b"\x89PNG\r\n\x1a\n" + parca(b"IHDR", struct.pack(">IIBBBBB", 24, 12, 8, 2, 0, 0, 0))
           + parca(b"IDAT", zlib.compress(ham)) + parca(b"IEND", b""))
    ns = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
    gorsel = ('<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="g"/>'
              '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
              '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
              '<a:blip r:embed="rId9"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic>'
              '</wp:inline></w:drawing></w:r>')
    hucre = lambda t: f'<w:tc><w:tcPr/><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:tc>'
    govde = ('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr>'
             '<w:t>SINAMA BELGESİ</w:t></w:r></w:p>'
             '<w:p><w:r><w:t xml:space="preserve">Türkçe karakterler: ğüşıöç İĞÜŞÖÇ ve </w:t></w:r>'
             '<w:r><w:rPr><w:i/></w:rPr><w:t>italik</w:t></w:r><w:r><w:t xml:space="preserve">, CO</w:t></w:r>'
             '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r></w:p>'
             '<w:tbl><w:tblPr><w:tblBorders><w:insideH w:val="single"/><w:insideV w:val="single"/>'
             '<w:top w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>'
             '<w:gridCol w:w="6000"/></w:tblGrid>'
             f'<w:tr>{hucre("Tarih")}{hucre("Olay")}</w:tr><w:tr>{hucre("01.01.2026")}{hucre("Tebliğ")}</w:tr>'
             f'</w:tbl><w:p>{gorsel}</w:p><w:p><w:r><w:t>Son paragraf.</w:t></w:r></w:p>')
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(yol, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/'
                   'package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.'
                   'openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType='
                   '"application/xml"/><Default Extension="png" ContentType="image/png"/><Override '
                   'PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
                   'wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.'
                   f'openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/'
                   'officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships '
                   f'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship '
                   f'Id="rId9" Type="{rel}/image" Target="media/g.png"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:document {ns}><w:body>'
                   f'{govde}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" '
                   'w:bottom="1417" w:left="1417"/></w:sectPr></w:body></w:document>')
        z.writestr("word/media/g.png", png)


def sinama(yollar=None, rapor_yolu=None):
    """Paketin bütünlüğünü doğrular (pencere göstermeden).

    Windows'ta --windowed derlenen exe'nin KONSOLU YOKTUR; print çıktısı hiçbir
    yere gitmez. Bu yüzden sonuç istenirse bir dosyaya da yazılır.
    """
    satirlar = []

    def yaz(*p):
        metin = " ".join(str(x) for x in p)
        satirlar.append(metin)
        try:
            print(metin)
        except UnicodeEncodeError:
            enc = sys.stdout.encoding or "ascii"
            print(metin.encode(enc, "replace").decode(enc, "replace"))
    kod = 0
    try:
        assert tk.TkVersion >= 8.6, f"Tcl/Tk {tk.TkVersion} çok eski (8.6+ gerekir)"
        klasor = tempfile.mkdtemp(prefix="udf_cevirici_sinama_")
        ornek = os.path.join(klasor, "sınama belgesi.docx")
        _ornek_docx(ornek)
        for y in [ornek] + list(yollar or []):
            r = belge_cevir(y, klasor)
            yaz(f"{'✓' if r['tamam'] else '✕'} {os.path.basename(y)}"
                + ("" if r["tamam"] else f" — {r['hata']}"))
            for u in r["uyarilar"]:
                yaz(f"    ! {u}")
            for d in r["dogrulama"]:
                yaz(f"    {d}")
            assert r["tamam"] or y != ornek, "örnek belge çevrilemedi: " + r["hata"]
        o = belge_cevir(ornek, klasor)["ozet"]
        assert o["tablo"] == 1 and o["paragraf"] >= 4, f"örnek belgenin yapısı beklenenden farklı: {o}"
        with zipfile.ZipFile(os.path.join(klasor, "sınama belgesi.udf")) as z:
            xml = z.read("content.xml").decode("utf-8")
        assert "<image " in xml and "<table " in xml and "ğüşıöç İĞÜŞÖÇ" in xml, "UDF içeriği eksik"
        kotu = belge_cevir(os.path.join(klasor, "yok.docx"), klasor)
        assert not kotu["tamam"] and kotu["hata"], "olmayan dosya için hata üretilmedi"

        # ters yön: UDF → Word. Asıl Word belgesi aynı klasörde duruyor; üzerine YAZILMAMALI.
        with open(ornek, "rb") as f:
            asil = f.read()
        t = belge_cevir(os.path.join(klasor, "sınama belgesi.udf"), klasor)
        yaz(f"{'✓' if t['tamam'] else '✕'} sınama belgesi.udf → {os.path.basename(t['cikti'])}"
            + ("" if t["tamam"] else f" — {t['hata']}"))
        for d in t["dogrulama"]:
            yaz(f"    {d}")
        assert t["tamam"], "UDF → Word çevirisi başarısız: " + t["hata"]
        assert os.path.basename(t["cikti"]) == "sınama belgesi (UDF'den).docx", "Word çıktısının adı: " + t["cikti"]
        with open(ornek, "rb") as f:
            assert f.read() == asil, "asıl Word belgesinin üzerine yazıldı"
        with zipfile.ZipFile(t["cikti"]) as z:
            govde = z.read("word/document.xml").decode("utf-8")
            assert any(n.startswith("word/media/") for n in z.namelist()), "Word çıktısında görsel yok"
        assert "<w:tbl>" in govde and "ğüşıöç İĞÜŞÖÇ" in govde, "Word çıktısının içeriği eksik"
        kotu = belge_cevir(os.path.join(klasor, "yok.udf"), klasor)
        assert not kotu["tamam"] and kotu["hata"], "olmayan UDF için hata üretilmedi"

        # Markdown → UDF yönü sınaması
        ornek_md = os.path.join(klasor, "sınama belgesi.md")
        with open(ornek_md, "w", encoding="utf-8") as f:
            f.write("# Sınama Belgesi (Markdown)\n\n"
                    "**ANKARA 1. ASLİYE HUKUK MAHKEMESİNE**\n\n"
                    "Bu bir *Markdown* belgesidir. ğüşıöç İĞÜŞÖÇ.\n\n"
                    "| Kalem | Miktar | Tutar |\n"
                    "|---|:---:|---:|\n"
                    "| Dava Değeri | 1 | 50.000 TL |\n\n"
                    "1. Birinci madde\n"
                    "2. İkinci madde\n")
        m = belge_cevir(ornek_md, klasor)
        yaz(f"{'✓' if m['tamam'] else '✕'} sınama belgesi.md → {os.path.basename(m['cikti'])}"
            + ("" if m["tamam"] else f" — {m['hata']}"))
        for d in m["dogrulama"]:
            yaz(f"    {d}")
        assert m["tamam"], "Markdown → UDF çevirisi başarısız: " + m["hata"]
        with zipfile.ZipFile(m["cikti"]) as z:
            m_xml = z.read("content.xml").decode("utf-8")
        assert "<table " in m_xml and "ğüşıöç İĞÜŞÖÇ" in m_xml, "Markdown UDF içeriği eksik"

        # Arayüz kuruluyor ve bir belgeyi uçtan uca çevirebiliyor mu? (pencere gösterilmez)
        uyg = Uygulama()
        uyg.withdraw()
        uyg.ekle([ornek])
        assert uyg.yollar == [os.path.abspath(ornek)], "belge listeye eklenemedi"
        uyg._cevir_bitti([belge_cevir(ornek, klasor)])
        uyg.update()
        assert uyg.sonuclar and uyg.sonuclar[0]["tamam"], "arayüz üzerinden çeviri başarısız"
        uyg.ekle([os.path.join(klasor, "sınama belgesi.udf")])
        assert len(uyg.yollar) == 2, "UDF belgesi listeye eklenemedi"
        ornek_uyari = [{"uyarilar": ["dipnotlar belge sonuna numaralı paragraf olarak taşındı (…)",
                                     "bu UDF e-imzalı; üretilen Word belgesi imza taşımaz (çalışma kopyasıdır)",
                                     "bilinmeyen yeni bir uyarı"]}]
        sade, kalan = sade_uyarilar(ornek_uyari)
        assert sade[0].startswith("Dipnotlar sayfa altında") and sade[1].startswith("Bu UDF e-imzalıydı;") \
            and sade[2] == "Bilinmeyen yeni bir uyarı" and kalan == 0, f"uyarı sadeleştirme bozuk: {sade}"
        dnd = "yok"
        if SURUKLENEBILIR:
            try:
                uyg.drop_target_register(DND_FILES)
                dnd = "çalışıyor"
            except Exception as e:
                dnd = f"KURULAMADI ({type(e).__name__})"
        editor = "bulundu" if uyg.editor_var else "yok"
        uyg.destroy()
        if editor == "bulundu":                              # paketteki Java sınıfları yerinde mi?
            try:
                _, sayfalar = udf_onizle.onizle(os.path.join(klasor, "sınama belgesi.udf"),
                                                os.path.join(klasor, "onizleme"))
                editor += f", önizleme çalışıyor ({len(sayfalar)} sayfa)"
            except Exception as e:
                editor += f", önizleme ÇALIŞMADI ({e})"
        assert ofis_docx.destekli("x.doc") and ofis_docx.destekli("x.pages") and not ofis_docx.destekli("x.pdf"), \
            "ofis biçimleri tanınmıyor"
        if sys.platform == "darwin":
            ofis = ", ".join(f"{ad}: {'var' if ofis_docx._mac_uygulama_var(b) else 'yok'}"
                             for ad, b in (("Word", ofis_docx.WORD_MAC), ("Pages", ofis_docx.PAGES_MAC)))
        else:
            ofis = "Word: çeviride denenecek"
        try:
            import PIL                                       # noqa
            pil = "Pillow " + PIL.__version__
        except Exception:
            pil = "Pillow yok (JPEG için sistem aracı denenir)"
        yaz(f"SINAMA TAMAM — iki yön, çekirdek ve arayüz çalışıyor (Tcl/Tk {tk.TkVersion}, "
            f"sürükle-bırak: {dnd}, {pil}, {ofis}, UYAP Editör: {editor}).")
    except AssertionError as e:
        yaz(f"SINAMA BAŞARISIZ — {e}")
        kod = 1
    except Exception as e:
        yaz(f"SINAMA BAŞARISIZ — beklenmeyen hata: {type(e).__name__}: {e}")
        kod = 1
    if rapor_yolu:
        try:
            with open(rapor_yolu, "w", encoding="utf-8") as f:
                f.write("\n".join(satirlar) + "\n")
        except OSError as e:
            print("rapor yazılamadı:", e, file=sys.stderr)
    return kod


def main():
    if "--sinama" in sys.argv:
        rapor = None
        if "--rapor" in sys.argv:
            i = sys.argv.index("--rapor")
            rapor = sys.argv[i + 1] if i + 1 < len(sys.argv) else None
        return sinama([y for y in sys.argv[1:] if y.lower().rstrip("/").endswith(KABUL)], rapor)
    if tk.TkVersion < 8.6:
        print(f"UYARI: Tcl/Tk {tk.TkVersion} çok eski; pencere boş görünebilir.", file=sys.stderr)
    uyg = Uygulama()
    uyg.ekle([y for y in sys.argv[1:] if os.path.exists(y)])
    uyg.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
