#!/usr/bin/env python3
"""
Modern UI Bileşenleri ve Etkileşim Testleri
"""
import os
import sys
import tempfile
import tkinter as tk
from tkinter import ttk
import unittest

from uygulama import (
    Uygulama, ModernListe, PALETLER, pencereyi_ortala,
    RaporPenceresi, SURUKLENEBILIR, _ornek_docx, belge_cevir
)

class TestModernUI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Uygulama()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    def setUp(self):
        self.root.temizle()
        self.root.sifirla()
        self.root.update_idletasks()

    def test_01_modern_liste_api(self):
        """ModernListe ttk.Treeview'un tk.Listbox arayüzünü eksiksiz taklit ettiğini doğrula"""
        liste = self.root.liste
        liste.delete(0, "end")
        self.assertEqual(liste.size(), 0)

        # Ekleme
        liste.ekle_oge("📄", "belge1.docx", "Word → UDF", tag="bekliyor")
        liste.ekle_oge("📄", "belge2.docx", "Word → UDF", tag="bekliyor")
        liste.ekle_oge("📄", "belge3.docx", "Word → UDF", tag="bekliyor")
        self.assertEqual(liste.size(), 3)

        # get tek ve aralık
        self.assertEqual(liste.get(0), "belge1.docx")
        self.assertEqual(liste.get(0, 1), ("belge1.docx", "belge2.docx"))

        # see ve selection_set
        liste.see(1)
        liste.selection_set(1)
        self.assertEqual(liste.curselection(), (1,))

        liste.selection_set(0, 2)
        self.assertEqual(liste.curselection(), (0, 1, 2))

        # selection_clear
        liste.selection_clear(1)
        self.assertEqual(liste.curselection(), (0, 2))
        liste.selection_clear(0, "end")
        self.assertEqual(liste.curselection(), ())

        # delete tek indeks
        liste.delete(1)
        self.assertEqual(liste.size(), 2)
        self.assertEqual(liste.get(0), "belge1.docx")
        self.assertEqual(liste.get(1), "belge3.docx")

        # delete aralık
        liste.delete(0, 1)
        self.assertEqual(liste.size(), 0)

    def test_02_tema_degisimi(self):
        """Açık ve koyu temalar arasında sorunsuz geçiş yapıldığını doğrula"""
        baslangic_koyu = self.root.tema_koyu
        self.root.tema_degistir()
        self.assertEqual(self.root.tema_koyu, not baslangic_koyu)
        self.root.update_idletasks()

        # Renk paleti kontrolleri
        tema_adi = "koyu" if self.root.tema_koyu else "acik"
        beklenen_bg = PALETLER[tema_adi]["bg_pencere"]
        self.assertEqual(self.root.cget("bg"), beklenen_bg)

        # Tekrar eski temaya dön
        self.root.tema_degistir()
        self.assertEqual(self.root.tema_koyu, baslangic_koyu)
        self.root.update_idletasks()

    def test_03_dinamik_dropzone_ve_kart_gorunurlugu(self):
        """Dosya yokken dropzone geniş, dosya varken kompakt olmalı"""
        # Liste boşken
        self.assertFalse(self.root._dropzone_kompakt)
        self.assertFalse(bool(self.root.liste_cerceve.winfo_manager()))
        self.assertFalse(bool(self.root.btn_cevir.winfo_manager()))

        # Geçici dosya oluşturup ekleyelim
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            gecici_docx = f.name

        try:
            self.root.ekle([gecici_docx])
            self.root.update_idletasks()

            # Dosya eklenince
            self.assertTrue(self.root._dropzone_kompakt)
            self.assertTrue(bool(self.root.liste_cerceve.winfo_manager()))
            self.assertTrue(bool(self.root.btn_cevir.winfo_manager()))
            self.assertEqual(len(self.root.yollar), 1)

            # Silme işlemi
            self.root.liste.selection_set(0)
            self.root.sil()
            self.root.update_idletasks()

            # Liste boşalınca
            self.assertFalse(self.root._dropzone_kompakt)
            self.assertFalse(bool(self.root.liste_cerceve.winfo_manager()))
            self.assertFalse(bool(self.root.btn_cevir.winfo_manager()))
            self.assertEqual(len(self.root.yollar), 0)
        finally:
            if os.path.exists(gecici_docx):
                os.remove(gecici_docx)

    def test_04_hedef_klasor_secimi_ve_sifirlama(self):
        """Hedef klasör atama ve sıfırlama işlevini doğrula"""
        self.root.hedef_klasor = "C:/TestKlasor"
        self.root.hedef_yazi.configure(text="C:/TestKlasor")
        self.root.btn_hedef_sifirla.pack(side="left")

        self.root.hedef_sifirla()
        self.assertIsNone(self.root.hedef_klasor)
        self.assertEqual(self.root.hedef_yazi.cget("text"), "📁 Belgenin yanına")
        self.assertFalse(bool(self.root.btn_hedef_sifirla.winfo_manager()))

    def test_05_gercek_zamanli_adim_guncellemeleri(self):
        """Çeviri adımlarında satır durumu ve ikonunun anlık güncellenmesi"""
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            d1 = f.name
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            d2 = f.name

        try:
            self.root.ekle([d1, d2])
            self.root.update_idletasks()

            # Adım 1 başladı
            self.root._adim_basladi(d1)
            item1 = self.root.liste.get_children()[0]
            degerler1 = self.root.liste.item(item1, "values")
            self.assertIn("Çevriliyor", degerler1[0])

            # Adım 1 bitti (başarılı)
            sonuc1 = {
                "kaynak": d1,
                "cikti": d1.replace(".docx", ".udf"),
                "tamam": True,
                "yon": "word>udf",
                "hata": None,
                "uyarilar": [],
                "dogrulama": ["OK"]
            }
            self.root._adim_bitti(sonuc1)
            degerler1 = self.root.liste.item(item1, "values")
            self.assertEqual(degerler1[0], "✓ Çevrildi")

            # Adım 2 başladı
            self.root._adim_basladi(d2)
            item2 = self.root.liste.get_children()[1]
            degerler2 = self.root.liste.item(item2, "values")
            self.assertIn("Çevriliyor", degerler2[0])

            # Adım 2 bitti (hatalı)
            sonuc2 = {
                "kaynak": d2,
                "cikti": "",
                "tamam": False,
                "yon": "word>udf",
                "hata": "Bozuk dosya",
                "uyarilar": [],
                "dogrulama": []
            }
            self.root._adim_bitti(sonuc2)
            degerler2 = self.root.liste.item(item2, "values")
            self.assertEqual(degerler2[0], "✕ Hata")

            # Çeviri bitti genel rapor
            self.root._cevir_bitti([sonuc1, sonuc2])
            self.root.update_idletasks()
            self.assertTrue(bool(self.root.durum_kutu.winfo_manager()))
            self.assertTrue(bool(self.root.sonuc_cerceve.winfo_manager()))

            # _secili_sonuc testi: hatalı dosya seçildiğinde hatalı sonucun referans alınması
            self.root.liste.selection_set(1)
            secili = self.root._secili_sonuc()
            self.assertIsNotNone(secili)
            self.assertEqual(secili["kaynak"], d2)
            self.assertFalse(secili["tamam"])

        finally:
            if os.path.exists(d1):
                os.remove(d1)
            if os.path.exists(d2):
                os.remove(d2)

    def test_06_diyaloglar_ve_pencere_ortala(self):
        """Rapor ve Hakkında pencerelerinin açılması ve ortalanması"""
        # Hakkında diyaloğu
        self.root.hakkinda_ac()
        self.root.update_idletasks()
        # Açılan toplevel'i bul ve kapat
        toplevels = [w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel)]
        self.assertTrue(len(toplevels) > 0)
        for t in toplevels:
            t.destroy()

        # Rapor penceresi
        rp = RaporPenceresi(self.root, "Test Rapor", "Deneme Raporu Metni", self.root.renk)
        self.root.update_idletasks()
        self.assertTrue(rp.winfo_exists())
        rp.destroy()

    def test_07_modern_liste_kenar_durumlari(self):
        """Boş liste, geçersiz indeks ve sınır durumlarının çökme üretmediğini doğrula"""
        liste = self.root.liste
        liste.delete(0, "end")

        # Boş listede işlemler
        self.assertEqual(liste.size(), 0)
        self.assertEqual(liste.get(0), "")
        self.assertEqual(liste.get("end"), "")
        self.assertEqual(liste.curselection(), ())
        liste.delete(0)
        liste.delete(0, 5)
        liste.delete("end")
        liste.see(0)
        liste.see("end")
        liste.selection_set(0)
        liste.selection_clear()
        liste.guncelle_oge(0, "test")

        # Öğeler ekleyip sınır aşımı testleri
        liste.ekle_oge("⏳", "a.docx", "Word → UDF")
        liste.ekle_oge("⏳", "b.docx", "Word → UDF")
        self.assertEqual(liste.size(), 2)

        # Sınır dışı get, see, delete
        self.assertEqual(liste.get(100), "")
        liste.see(100)
        liste.selection_set(100)
        self.assertEqual(liste.curselection(), ())
        liste.selection_clear(100)
        liste.delete(100)
        self.assertEqual(liste.size(), 2)

        # "end" seçimi ve silinmesi
        liste.selection_set("end")
        self.assertEqual(liste.curselection(), (1,))
        liste.delete("end")
        self.assertEqual(liste.size(), 1)
        self.assertEqual(liste.get(0), "a.docx")

    def test_08_dosya_filtreleme_ve_coklu_silme(self):
        """Geçersiz dosya, geçici dosya ve çoklu seçim silme akışlarını doğrula"""
        with tempfile.TemporaryDirectory() as td:
            f1 = os.path.join(td, "doc1.docx")
            f2 = os.path.join(td, "doc2.docx")
            f3 = os.path.join(td, "doc3.docx")
            f_temp = os.path.join(td, "~$temp.docx")
            f_pdf = os.path.join(td, "test.pdf")

            for f in (f1, f2, f3, f_temp, f_pdf):
                with open(f, "w", encoding="utf-8") as fp:
                    fp.write("dummy")

            # Ekleme: f1, f2, f3 kabul edilmeli; f_temp ve f_pdf atlanmalı
            self.root.ekle([f1, f2, f3, f_temp, f_pdf, "olmayan_dosya.docx"])
            self.root.update_idletasks()

            self.assertEqual(len(self.root.yollar), 3)
            self.assertEqual(self.root.liste.size(), 3)
            self.assertTrue(bool(self.root.durum_kutu.winfo_manager()))  # atlanan uyarısı çıkmalı

            # Çoklu seçim: 0. ve 2. indeksleri seçip sil
            self.root.liste.selection_set(0)
            self.root.liste.selection_set(2)
            self.assertEqual(self.root.liste.curselection(), (0, 2))

            self.root.sil()
            self.root.update_idletasks()

            # Yalnız f2 kalmalı
            self.assertEqual(len(self.root.yollar), 1)
            self.assertEqual(self.root.yollar[0], os.path.abspath(f2))
            self.assertEqual(self.root.liste.size(), 1)

            # Temizle çağrısı
            self.root.temizle()
            self.root.update_idletasks()
            self.assertEqual(len(self.root.yollar), 0)
            self.assertEqual(self.root.liste.size(), 0)
            self.assertFalse(bool(self.root.liste_cerceve.winfo_manager()))

    def test_09_kaydirilabilir_govde_ve_fare_tekeri(self):
        """Pencere kaydırılabilir gövde, scrollbar ve fare tekeri etkileşimlerini doğrula"""
        self.assertTrue(hasattr(self.root, "canvas"))
        self.assertTrue(hasattr(self.root, "scrollbar"))
        self.assertTrue(hasattr(self.root, "ana_tasiyici"))
        self.assertTrue(hasattr(self.root, "govde_tasiyici"))

        # İçeriğe widget ekleyip scrollregion ve görünürlük fonksiyonunu sına
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            d = f.name
        try:
            self.root.ekle([d])
            self.root.update_idletasks()

            # Dönüşüm ayarları ve buton görünürlüğü
            self.assertTrue(bool(self.root.ayar.winfo_manager()))
            self.assertTrue(bool(self.root.btn_cevir.winfo_manager()))

            # _gorunur_yap çağrısı
            self.root._gorunur_yap(self.root.btn_cevir)
            self.root.update_idletasks()

            # Fare tekeri simülasyonu (delta)
            event_mock = type("MockEvent", (), {
                "x_root": self.root.winfo_rootx() + 20,
                "y_root": self.root.winfo_rooty() + 20,
                "delta": -120
            })()
            ret = self.root._fare_tekeri(event_mock)
            self.assertEqual(ret, "break")
        finally:
            if os.path.exists(d):
                os.remove(d)

if __name__ == "__main__":
    unittest.main()

