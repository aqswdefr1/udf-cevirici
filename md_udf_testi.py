#!/usr/bin/env python3
r"""md_udf_testi.py — Markdown -> UDF dönüştürücüsünün kapsamlı senaryo testleri.

Test edilen başlıklar:
  1. Başlıklar (H1-H6, setext, mahkeme adı kendiliğinden ortalama)
  2. Karakter biçimleri (kalın, italik, üstü çizili, altı çizili, vurgu, üst/alt simge, kod)
  3. HTML & CSS etiketleri (<span style="..">, <font>, <mark>, <br> vb.)
  4. Kaçış karakterleri (\*, \_, \` vb.) ve snake_case güvenliği
  5. Paragraf hizalamaları (iki yana yasla, sola, sağa, ortaya, <center>, <p align="..">)
  6. Listeler (madde işaretli, numaralı, harfli, Roma rakamlı, iç içe listeler, başlama sayısı)
  7. Tablolar (hizalamalar, sütun oranları, boş hücreler, hücre içi boru kaçışı)
  8. Alıntılar (> tek düzey, >> iç içe)
  9. Kod blokları (``` fenced, Courier New yazı tipi, girinti ve boşluk koruma)
  10. Sayfa sonu ve yatay çizgiler (<page-break>, \pagebreak, ---sayfa sonu---, --- vb.)
  11. Dipnotlar ([^1] üst simge ve belge sonu dipnot bloğu)
  12. Görseller (data URI, yerel dosya, olmayan dosya için güvenli geri dönüş)
  13. Dilekçe şablonu (mahkeme başlığı, taraflar, açıklamalar, talep tablosu, imza bloğu)
  14. Türkçe karakterler, CDATA kaçışı (]]>) ve emoji / BMP dışı karakter güvenliği
  15. Hata durumları (olmayan dosya vb.)
  16. Tam döngü: Markdown -> UDF -> Word (DOCX) ve udf_dogrula denetimi

Kullanım:
    python3 md_udf_testi.py
"""

import os
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

import md_udf
import udf_docx
import udf_dogrula
from docx_udf import DonusumHatasi

HERE = os.path.dirname(os.path.abspath(__file__))


class MdUdfTestleri(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="md_udf_test_")

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _cevir_ve_oku(self, md_metin, **kwargs):
        """Markdown metnini UDF'ye çevirip (xml, cdata, ozet, udf_yolu) döndürür."""
        md_yolu = os.path.join(self.temp_dir, "test.md")
        udf_yolu = os.path.join(self.temp_dir, "test.udf")
        with open(md_yolu, "w", encoding="utf-8") as f:
            f.write(md_metin)

        ozet = md_udf.cevir(md_yolu, udf_yolu, **kwargs)
        with zipfile.ZipFile(udf_yolu) as z:
            xml_raw = z.read("content.xml").decode("utf-8")

        root = ET.fromstring(xml_raw.split("?>", 1)[1].strip())
        cdata = root.find("content").text or ""
        return root, cdata, ozet, md_yolu, udf_yolu

    # ----------------------------------------------------------------------
    # 1. Başlıklar
    # ----------------------------------------------------------------------
    def test_basliklar(self):
        md = (
            "# Birinci Başlık\n\n"
            "## İkinci Başlık\n\n"
            "### Üçüncü Başlık\n\n"
            "#### Dördüncü Başlık\n\n"
            "##### Beşinci Başlık\n\n"
            "###### Altıncı Başlık\n\n"
            "Setext H1\n"
            "===\n\n"
            "Setext H2\n"
            "---\n\n"
            "# İSTANBUL 1. ASLİYE TİCARET MAHKEMESİNE\n\n"
            "# T.C.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("Birinci Başlık", cdata)
        self.assertIn("İSTANBUL 1. ASLİYE TİCARET MAHKEMESİNE", cdata)

        # Mahkeme başlığı ve T.C. ortalanmış olmalı (Alignment="1")
        cd16 = cdata.encode('utf-16-le')
        paras = root.findall(".//paragraph")
        def _p_text(p):
            t_list = []
            for c in p.findall("content"):
                so = int(c.attrib.get("startOffset", 0))
                ln = int(c.attrib.get("length", 0))
                t_list.append(cd16[2 * so:2 * (so + ln)].decode('utf-16-le', 'replace'))
            return "".join(t_list)

        mahkeme_p = [p for p in paras if "İSTANBUL 1. ASLİYE" in _p_text(p)]
        self.assertTrue(len(mahkeme_p) > 0, "Mahkeme başlığı paragrafı bulunamadı")
        self.assertEqual(mahkeme_p[0].attrib.get("Alignment"), "1", "Mahkeme başlığı ortalanmamış")

        tc_p = [p for p in paras if "T.C." in _p_text(p)]
        self.assertTrue(len(tc_p) > 0, "T.C. başlığı paragrafı bulunamadı")
        self.assertEqual(tc_p[0].attrib.get("Alignment"), "1", "T.C. ortalanmamış")

        # Doğrulama: dogrula fonksiyonu GEÇTİ vermeli
        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 2. Karakter Biçimleri
    # ----------------------------------------------------------------------
    def test_karakter_bicimleri(self):
        md = (
            "Bu metinde **kalın yazı**, *italik yazı*, ***kalın italik***, "
            "~~üstü çizili~~, ==vurgulu yazı==, <u>altı çizili</u>, "
            "^üst simge^, ~alt simge~ ve `kod parçası` bulunmaktadır.\n\n"
            "Yılan_durumu_degiskeni_icin_italik_olmamali_guvenlik_testi.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("kalın yazı", cdata)
        self.assertIn("vurgulu yazı", cdata)
        self.assertIn("Yılan_durumu_degiskeni_icin_italik_olmamali_guvenlik_testi", cdata)

        # Vurgu rengi kontrolü (background / foreground)
        contents = root.findall(".//content")
        has_bold = any(c.attrib.get("bold") == "true" for c in contents)
        has_italic = any(c.attrib.get("italic") == "true" for c in contents)
        has_strike = any(c.attrib.get("strikethrough") == "true" for c in contents)
        has_underline = any(c.attrib.get("underline") == "true" for c in contents)
        has_bg = any(c.attrib.get("background") for c in contents)
        has_code_font = any(c.attrib.get("family") == "Courier New" for c in contents)

        self.assertTrue(has_bold, "Kalın biçim bulunamadı")
        self.assertTrue(has_italic, "İtalik biçim bulunamadı")
        self.assertTrue(has_strike, "Üstü çizili biçim bulunamadı")
        self.assertTrue(has_underline, "Altı çizili biçim bulunamadı")
        self.assertTrue(has_bg, "Vurgu arka planı bulunamadı")
        self.assertTrue(has_code_font, "Kod yazı tipi (Courier New) bulunamadı")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 3. HTML ve CSS Etiketleri
    # ----------------------------------------------------------------------
    def test_html_ve_css_etiketleri(self):
        md = (
            'Burada <span style="color: #FF0000; font-size: 14pt; font-family: Arial">Kırmızı 14pt Arial</span> var.\n\n'
            'Burada <span style="background-color: yellow">Sarı Zemin</span> ve '
            '<font color="blue" face="Calibri">Mavi Calibri</font> var.\n\n'
            'Satır 1<br>Satır 2<br/>Satır 3\n'
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("Kırmızı 14pt Arial", cdata)
        self.assertIn("Mavi Calibri", cdata)

        contents = root.findall(".//content")
        # Kırmızı renk foreground=-65536 olmalı
        has_red = any(c.attrib.get("foreground") == "-65536" for c in contents)
        has_arial = any(c.attrib.get("family") == "Arial" for c in contents)
        has_14pt = any(c.attrib.get("size") == "14" for c in contents)

        self.assertTrue(has_red, "Kırmızı renk bulunamadı")
        self.assertTrue(has_arial, "Arial yazı tipi bulunamadı")
        self.assertTrue(has_14pt, "14 punto boyutu bulunamadı")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 4. Kaçış Karakterleri
    # ----------------------------------------------------------------------
    def test_kacis_karakterleri(self):
        md = r"Burada \*yıldız\* ve \_alt çizgi\_ ve \`ters tırnak\` ve \[köşeli\] korunmalı."
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("*yıldız*", cdata)
        self.assertIn("_alt çizgi_", cdata)
        self.assertIn("`ters tırnak`", cdata)
        self.assertIn("[köşeli]", cdata)

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 5. Paragraf Hizalamaları
    # ----------------------------------------------------------------------
    def test_paragraf_hizalamalari(self):
        md = (
            "<center>Ortalanmış Paragraf</center>\n\n"
            '<p align="right">Sağa Yaslı Paragraf</p>\n\n'
            '<div align="left">Sola Yaslı Paragraf</div>\n\n'
            '<p align="justify">İki Yana Yaslı Paragraf</p>\n\n'
            "-> İşaretle Ortalanmış Paragraf <-\n\n"
            "Varsayılan Hizada Paragraf\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md, hiza="ikiye-yasla")
        paras = root.findall(".//paragraph")
        aligns = [p.attrib.get("Alignment") for p in paras]
        self.assertIn("1", aligns, "Ortalanmış paragraf (Alignment=1) bulunamadı")
        self.assertIn("2", aligns, "Sağa yaslı paragraf (Alignment=2) bulunamadı")
        self.assertIn("0", aligns, "Sola yaslı paragraf (Alignment=0) bulunamadı")
        self.assertIn("3", aligns, "İki yana yaslı paragraf (Alignment=3) bulunamadı")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 6. Listeler
    # ----------------------------------------------------------------------
    def test_listeler(self):
        md = (
            "- Birinci madde\n"
            "- İkinci madde\n"
            "  * Alt madde A\n"
            "  * Alt madde B\n"
            "    + Üçüncü düzey\n\n"
            "1. Numaralı bir\n"
            "2. Numaralı iki\n\n"
            "a. Harfli bent\n"
            "b. İkinci bent\n\n"
            "5. Beşten başlayan madde\n"
            "6. Altıncı madde\n\n"
            "I. Roma birinci\n"
            "II. Roma ikinci\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("Birinci madde", cdata)
        self.assertIn("Alt madde A", cdata)
        self.assertIn("Numaralı bir", cdata)
        self.assertIn("Beşten başlayan madde", cdata)

        # Liste öznitelikleri kontrolü
        paras = root.findall(".//paragraph")
        bulleted = [p for p in paras if p.attrib.get("Bulleted") == "true"]
        numbered = [p for p in paras if p.attrib.get("Numbered") == "true"]

        self.assertTrue(len(bulleted) >= 3, "Madde işaretli liste öğeleri eksik")
        self.assertTrue(len(numbered) >= 6, "Numaralı liste öğeleri eksik")

        # 5'ten başlayan listede NumberSetted="4" olmalı (n+1 kuralı)
        has_setted_4 = any(p.attrib.get("NumberSetted") == "4" for p in numbered)
        self.assertTrue(has_setted_4, "NumberSetted='4' bulunamadı")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 7. Tablolar
    # ----------------------------------------------------------------------
    def test_tablolar(self):
        md = (
            "| Sıra | Ad Soyad | Açıklama | Tutar |\n"
            "|:---:|:---|---|---:|\n"
            "| 1 | Ahmet Yılmaz | Peşinat | 10.000 TL |\n"
            "| 2 | Mehmet Demir | Kalan Borç | 25.500 TL |\n"
            "| 3 | Ayşe Kaya | `Kod \\| Borç` | 5.000 TL |\n"
            "| 4 | | Boş Hücre Testi | 0 TL |\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertEqual(ozet["tablo"], 1)
        self.assertIn("Ahmet Yılmaz", cdata)
        self.assertIn("Mehmet Demir", cdata)

        tables = root.findall(".//table")
        self.assertEqual(len(tables), 1)
        tb = tables[0]
        self.assertEqual(tb.attrib.get("columnCount"), "4")
        spans = [int(x) for x in tb.attrib.get("columnSpans").split(",")]
        self.assertEqual(sum(spans), 400, "Sütun oranları toplamı 400 (100 * ncol) olmalıdır")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 8. Alıntılar
    # ----------------------------------------------------------------------
    def test_alintilar(self):
        md = (
            "> Yargıtay 11. Hukuk Dairesi'nin yerleşik içtihatlarına göre:\n"
            "> \"Tacirler basiretli bir iş insanı gibi davranmakla yükümlüdür.\"\n"
            ">\n"
            ">> İç içe alıntı örneği ikinci düzeyde yer alır.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("Yargıtay 11. Hukuk Dairesi", cdata)
        self.assertIn("İç içe alıntı örneği", cdata)

        paras = root.findall(".//paragraph")
        has_left_indent = any(float(p.attrib.get("LeftIndent", 0)) >= 36.0 for p in paras)
        self.assertTrue(has_left_indent, "Alıntı girintisi (LeftIndent >= 36) bulunamadı")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 9. Kod Blokları
    # ----------------------------------------------------------------------
    def test_kod_bloklari(self):
        md = (
            "Aşağıdaki kod fonksiyonu hesaplama yapar:\n\n"
            "```python\n"
            "def tazminat_hesapla(faiz, tutar):\n"
            "    toplam = tutar * (1 + faiz)\n"
            "    return toplam\n"
            "```\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("def tazminat_hesapla", cdata)
        self.assertIn("return toplam", cdata)

        contents = root.findall(".//content")
        has_courier = any(c.attrib.get("family") == "Courier New" for c in contents)
        self.assertTrue(has_courier, "Kod bloğu Courier New yazı tipi içermiyor")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 10. Sayfa Sonu ve Bölücüler
    # ----------------------------------------------------------------------
    def test_sayfa_sonu_ve_boluculer(self):
        md = (
            "Sayfa 1 Metni\n\n"
            "---sayfa sonu---\n\n"
            "Sayfa 2 Metni\n\n"
            "<page-break/>\n\n"
            "Sayfa 3 Metni\n\n"
            "---\n\n"
            "Yatay Çizgi Sonrası Metin\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertEqual(ozet["sayfa_sonu"], 2, "2 adet sayfa sonu bekleniyordu")
        pbs = root.findall(".//page-break")
        self.assertEqual(len(pbs), 2)

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 11. Dipnotlar
    # ----------------------------------------------------------------------
    def test_dipnotlar(self):
        md = (
            "Davacının talebi haksız fiile dayanmaktadır[^1]. "
            "Ayrıca TBK madde 49 uyarınca tazminat istenmiştir[^kanun].\n\n"
            "[^1]: Yargıtay Hukuk Genel Kurulu 2024/100 E.\n"
            "[^kanun]: 6098 sayılı Türk Borçlar Kanunu.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("haksız fiile dayanmaktadır", cdata)
        self.assertIn("Yargıtay Hukuk Genel Kurulu", cdata)

        # Üst simge dipnot işareti kontrolü
        contents = root.findall(".//content")
        has_fn_super = any(c.attrib.get("superscript") == "true" for c in contents)
        self.assertTrue(has_fn_super, "Dipnot üst simgesi bulunamadı")

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 12. Görseller
    # ----------------------------------------------------------------------
    def test_gorseller(self):
        # 1x1 piksellik geçerli bir PNG görseli base64 data URI olarak
        # Red pixel PNG: 1x1
        png_1x1 = (
            b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01'
            b'\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0'
            b'\x00\x00\x03\x01\x01\x00\x18\xdd\x8d\xb0\x00\x00\x00\x00IEND\xaeB`\x82'
        )
        import base64
        b64 = base64.b64encode(png_1x1).decode('ascii')
        data_uri = f"data:image/png;base64,{b64}"

        # Yerel dosya oluştur
        yerel_png = os.path.join(self.temp_dir, "logo.png")
        with open(yerel_png, "wb") as f:
            f.write(png_1x1)

        md = (
            f"İşte bir base64 görsel: ![Kırmızı Piksel]({data_uri})\n\n"
            f"İşte yerel bir görsel: ![Yerel Logo](logo.png)\n\n"
            f"Olmayan görsel için geri dönüş: ![Olmayan](yok_resim.png)\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertEqual(ozet["gorsel"], 2, "2 geçerli görsel bulunmalıydı")
        self.assertTrue(any("bulunamadı" in u or "okunamadı" in u for u in ozet["uyarilar"]))

        images = root.findall(".//image")
        self.assertEqual(len(images), 2)

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 13. Kapsamlı Dava Dilekçesi Şablonu
    # ----------------------------------------------------------------------
    def test_dilekce_sablonu(self):
        md = (
            "---\n"
            "title: TAZMİNAT DAVA DİLEKÇESİ\n"
            "author: Av. Ahmet Yılmaz\n"
            "date: 2026-10-07\n"
            "---\n\n"
            "# ANKARA 2. ASLİYE HUKUK MAHKEMESİNE\n\n"
            "**DAVACI\t:** Mehmet Demir (T.C. Kimlik No: 12345678901)\n\n"
            "**VEKİLİ\t:** Av. Ahmet Yılmaz, Ankara Barosu No: 99999\n\n"
            "**DAVALI\t:** Anadolu Sigorta A.Ş.\n\n"
            "**DAVA KONUSU\t:** Trafik kazası neticesinde maddi tazminat istemidir.\n\n"
            "**DAVA DEĞERİ\t:** 100.000,00 TL\n\n"
            "## AÇIKLAMALAR\n\n"
            "1. Müvekkil sevk ve idaresindeki araç, davalı nezdinde ZMMS poliçesi ile sigortalıdır.\n"
            "2. Meydana gelen kazada karşı araç sürücüsü %100 kusurludur.\n"
            "3. Hasar tespit raporuna göre araçta pert-total durumu oluşmuştur.\n\n"
            "### ZARAR KALEMLERİ TABLOSU\n\n"
            "| Sıra | Zarar Türü | Talep Tutarı | Yasal Faiz Başlangıcı |\n"
            "|:---:|:---|---:|:---:|\n"
            "| 1 | Araç Değer Kaybı | 40.000,00 TL | 12/03/2026 |\n"
            "| 2 | Bakiye Araç Bedeli | 60.000,00 TL | 12/03/2026 |\n"
            "| | **TOPLAM** | **100.000,00 TL** | |\n\n"
            "## HUKUKİ DELİLLER\n\n"
            "- Kaza tespit tutanağı\n"
            "- Sigorta poliçesi\n"
            "- Bilirkişi incelemesi\n\n"
            "## SONUÇ VE İSTEM\n\n"
            "Yukarıda arz ve izah edilen nedenlerle davanın kabulüne karar verilmesini arz ve talep ederiz.\n\n"
            '<p align="right">Davacı Vekili<br>Av. Ahmet Yılmaz<br>(e-imzalıdır)</p>\n'
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("TAZMİNAT DAVA DİLEKÇESİ", cdata)
        self.assertIn("ANKARA 2. ASLİYE HUKUK MAHKEMESİNE", cdata)
        self.assertIn("ZARAR KALEMLERİ TABLOSU", cdata)
        self.assertIn("Araç Değer Kaybı", cdata)
        self.assertIn("Davacı Vekili", cdata)

        # Doğrulama: hem udf_dogrula hem XML ofset denetimi geçmeli
        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 14. Türkçe Karakterler, CDATA ve Emoji Güvenliği
    # ----------------------------------------------------------------------
    def test_turkce_ve_emoji_guvenligi(self):
        md = (
            "Türkçe karakterler: ğüşıöç İĞÜŞÖÇ, şapkalı harfler: kâğıt, resmî, dâhil.\n\n"
            "CDATA kaçış testi: metin içinde ]] > ya da ]]> olması dosyayı bozmamalıdır.\n\n"
            "Emoji testi: Adalet terazisi ⚖️, saray 🏛️, tebessüm 😊 simgeleri Editör uyumu için □ simgesine dönüşmeli.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("ğüşıöç İĞÜŞÖÇ", cdata)
        self.assertIn("kâğıt", cdata)
        self.assertNotIn("]]>", cdata, "CDATA kapatma etiketi ']]>' zararsızlaştırılmalı")

        # Emoji uyarısı üretilmiş olmalı
        self.assertTrue(any("Editör'ün açamadığı karakter" in u for u in ozet["uyarilar"]))

        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    # ----------------------------------------------------------------------
    # 15. Hata Durumları
    # ----------------------------------------------------------------------
    def test_hata_durumlari(self):
        olmayan = os.path.join(self.temp_dir, "olmayan_dosya.md")
        cikti = os.path.join(self.temp_dir, "cikti.udf")
        # Olmayan dosya dönüştürülmek istendiğinde temiz DonusumHatasi atılmalı
        with self.assertRaises(DonusumHatasi):
            md_udf.cevir(olmayan, cikti)

    # ----------------------------------------------------------------------
    # 16. Tam Döngü: Markdown -> UDF -> Word (DOCX)
    # ----------------------------------------------------------------------
    def test_md_udf_docx_donusumu(self):
        md = (
            "# GENEL KURUL KARARI\n\n"
            "**Tarih:** 07.10.2026\n\n"
            "Aşağıdaki gündem maddeleri karara bağlanmıştır:\n\n"
            "1. Yönetim kurulu faaliyet raporu okundu ve onaylandı.\n"
            "2. Denetçi raporu ibra edildi.\n\n"
            "| Madde | Karar | Oy Sayısı |\n"
            "|---|---|---:|\n"
            "| 1 | Kabul | 100 |\n"
            "| 2 | Kabul | 98 |\n\n"
            "Toplantı sona erdi.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)

        # UDF doğrulamasından geçmeli
        ok_udf, satirlar_udf = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok_udf, "\n".join(satirlar_udf))

        # Ters yön: UDF -> DOCX
        docx_yolu = os.path.join(self.temp_dir, "test.docx")
        udf_docx.cevir(udf_yolu, docx_yolu)
        self.assertTrue(os.path.exists(docx_yolu))

        # DOCX doğrulamasından geçmeli
        ok_docx, satirlar_docx = udf_dogrula.dogrula_docx(udf_yolu, docx_yolu)
        self.assertTrue(ok_docx, "\n".join(satirlar_docx))

        # DOCX içeriğinde metinler ve tablo XML'i bulunmalı
        with zipfile.ZipFile(docx_yolu) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
        self.assertIn("GENEL KURUL KARARI", doc_xml)
        self.assertIn("<w:tbl>", doc_xml)
        self.assertIn("Yönetim kurulu", doc_xml)

    # ----------------------------------------------------------------------
    # 17. Uç Durumlar ve Sağlamlık Testleri
    # ----------------------------------------------------------------------
    def test_kelime_yutulmama_aciklamalar(self):
        """'Açıklamalar.' gibi satır başı kelimeler liste işareti sanılıp silinmemeli."""
        md = (
            "Açıklamalar. Davalı şirket borcunu süresinde ödememiştir.\n\n"
            "Özet. Bu bir hukuki özet metnidir.\n\n"
            "Hüküm. Davanın kabulüne karar verildi.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("Açıklamalar.", cdata)
        self.assertIn("Özet.", cdata)
        self.assertIn("Hüküm.", cdata)
        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    def test_bos_belge_ve_yatay_cizgi(self):
        """Boş belge ve frontmatter olmayan '---' çizgisi güvenle işlenmeli."""
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku("")
        self.assertGreaterEqual(ozet["paragraf"], 1)

        md_hr = "---\n\nİlk Paragraf\n\n---\n"
        root2, cdata2, ozet2, md_yolu2, udf_yolu2 = self._cevir_ve_oku(md_hr)
        self.assertIn("İlk Paragraf", cdata2)
        ok, satirlar = udf_dogrula.dogrula(md_yolu2, udf_yolu2)
        self.assertTrue(ok, "\n".join(satirlar))

    def test_alt_simge_unicode(self):
        """Alt simgeler (H~2~O) Editör için Unicode SUBSCRIPT karakterine çevrilmeli."""
        md = "Kimyasal formül: H~2~O ve CO~2~ gazı."
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("H₂O", cdata)
        self.assertIn("CO₂", cdata)
        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    def test_harfli_liste_alfabe_uyarisi(self):
        """4 ve daha fazla kalemli harfli listelerde Türk alfabesi uyarısı verilmeli."""
        md = (
            "a. Birinci madde\n"
            "b. İkinci madde\n"
            "c. Üçüncü madde\n"
            "d. Dördüncü madde\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertTrue(any("Türk alfabesi" in u for u in ozet["uyarilar"]),
                        f"Harfli liste uyarısı eksik: {ozet['uyarilar']}")

    def test_cift_parantez_liste(self):
        """(1) ve (a) şeklindeki fıkra/bent listeleri D_PARANTHESE olarak tanınmalı."""
        md = (
            "(1) Birinci fıkra metni.\n"
            "(2) İkinci fıkra metni.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("Birinci fıkra metni.", cdata)
        paras = root.findall(".//paragraph")
        has_d_paren = any("D_PARANTHESE" in p.attrib.get("NumberType", "") for p in paras)
        self.assertTrue(has_d_paren, "D_PARANTHESE liste türü bulunamadı")
        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    def test_atifsiz_dipnot(self):
        """Gövde metninde atıf yapılmayan dipnot tanımları da belge sonuna eklenmeli."""
        md = (
            "Bu metinde doğrudan atıf yok.\n\n"
            "[^tanim]: Bu dipnot metin içinde çağrılmadı ama korunmalı.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("Bu dipnot metin içinde çağrılmadı ama korunmalı.", cdata)
        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))

    def test_ardisik_tablolar_ve_sonlandirma(self):
        """Ardışık tablolar arasına ve belge sonuna paragraf eklenmeli (blok normalizasyonu)."""
        md = (
            "| A | B |\n"
            "|---|---|\n"
            "| 1 | 2 |\n\n"
            "| C | D |\n"
            "|---|---|\n"
            "| 3 | 4 |\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        els = root.find("elements")
        body_children = [el for el in list(els) if el.tag not in ("header", "footer")]
        tbl_indices = [idx for idx, el in enumerate(body_children) if el.tag == "table"]
        self.assertEqual(len(tbl_indices), 2)
        self.assertGreater(tbl_indices[1] - tbl_indices[0], 1, "İki tablo arasında paragraf yok")
        self.assertEqual(body_children[-1].tag, "paragraph", "Belge gövde sonu paragraf ile bitmiyor")

    def test_cok_satirli_dipnot_tanimi(self):
        """Girintili çok satırlı dipnot tanımları tek dipnotta toplanmalı."""
        md = (
            "Metin burada[^not1].\n\n"
            "[^not1]: İlk açıklama satırı.\n"
            "    İkinci satır devamı.\n"
        )
        root, cdata, ozet, md_yolu, udf_yolu = self._cevir_ve_oku(md)
        self.assertIn("İlk açıklama satırı.", cdata)
        self.assertIn("İkinci satır devamı.", cdata)
        ok, satirlar = udf_dogrula.dogrula(md_yolu, udf_yolu)
        self.assertTrue(ok, "\n".join(satirlar))


if __name__ == "__main__":
    unittest.main()

