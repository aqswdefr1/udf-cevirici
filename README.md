# UDF Çevirici

Word ve Markdown belgelerini, biçimini koruyarak UYAP Doküman Editörü'nün **UDF** biçimine, UDF belgesini
de **Word**'e çeviren küçük bir masaüstü programı. `.docx` ve `.md` dışında `.doc`, `.rtf`, `.odt` ve (Mac'te)
`.pages` belgeleri de çevrilir; bunlar önce bilgisayardaki Word'e ya da Pages'e Word biçimine çevirtilir. Yönü dosyanın uzantısı belirler. Windows ve
macOS'ta çalışır; kurulum gerektirmez.

- Tablolar (birleştirilmiş hücreler dâhil), görseller, numaralı ve madde işaretli listeler,
  üstbilgi/altbilgi ve logo, sayfa sonları, dipnotlar, sekme durakları, yazı biçimleri taşınır.
- Her çıktı üretildikten sonra kaynağıyla **karşılaştırılarak doğrulanır**; doğrulama
  geçmezse dosya yazılmaz.
- UDF → Word yönünde var olan bir Word belgesinin üzerine hiçbir zaman yazılmaz
  (`Dilekçe (UDF'den).docx`); görseller yeniden kodlanmadan, bayt bayt aktarılır.
- Bilgisayarda UYAP Doküman Editörü kuruluysa, Editör'ü açmadan **Editör'ün kendi çizim
  motoruyla önizleme** gösterir.
- Belgeler bilgisayardan çıkmaz; program kendiliğinden internete bağlanmaz.

## İndir

**[Son sürümü indirin →](https://github.com/miasimbilir/udf-cevirici/releases/latest)**

| Bilgisayar | Dosya |
|---|---|
| Windows 10/11 | `UDF-Cevirici-…-Windows.zip` |
| Mac (Apple M işlemcili) | `UDF-Cevirici-…-macOS-AppleSilicon.zip` |
| Mac (Intel işlemcili) | `UDF-Cevirici-…-macOS-Intel.zip` |

Zip'i açın, programı çalıştırın; kurulum gerekmez. Program imzasız olduğu için ilk açılışta uyarı verir:
Windows'ta **"Ek bilgi" → "Yine de çalıştır"**, Mac'te uygulamaya **sağ tıklayıp "Aç"**.
Programın içindeki **Hakkında → Güncellemeleri kontrol et** düğmesi yeni sürüm olup olmadığını söyler
(program yalnızca bu düğmeye basınca internete bağlanır, belge göndermez).

Kullanım için: [KULLANIM.txt](KULLANIM.txt) · Derleme için: [DERLEME.md](DERLEME.md)

## Neden gerekli

UYAP Doküman Editörü Word belgesini içe aktarırken tabloları, girintileri ve listeleri çoğu
zaman bozar. Bu program DOCX'in yapısını doğrudan okuyup UDF'yi Editör'ün kendi ürettiği
biçimde yazar.

## UDF'nin sınırları ve programın taklitleri

UDF, Word'ün her özelliğini karşılamaz. Aşağıdakiler Editör'ün kendi dosyaları ve çizim
motoru üzerinde sınanarak belirlenmiştir:

| Word özelliği | UDF'de durum | Programın yaptığı |
|---|---|---|
| Hücre dolgu rengi | yok | yazının arkasını hücre rengiyle boyar, bandı satır boyunca uzatır |
| Hücrede dikey hizalama, iç boşluk | yok | paragraf boşluklarıyla taklit eder (yükseklik tahminidir) |
| Dikey hücre birleştirme | yok | iç içe tabloyla kurar; karmaşık durumda alt hücre boş kalır |
| Yatay hücre birleştirme | var | satır düzeyinde sütun oranlarıyla yazar |
| Sayfa altı dipnot | yok | işareti üst simge yapar, metni belge sonuna koyar |
| İlk sayfa + devam için iki üstbilgi | tek üstbilgi | ilk sayfanınkini korur |
| Metin kaydırmalı görsel | yok | satır içine alır (deneysel taklit: `--yan-gorsel`) |
| JPEG/GIF/BMP görsel | yalnız PNG | PNG'ye çevirir |
| Emoji vb. (BMP dışı) karakter | Editör dosyayı açamaz | □ ile değiştirir |
| Alt simge | Editör yukarıda çizer | Unicode alt simge karakterine çevirir (CO₂) |
| Sekme lider noktaları, şekil, grafik, EMF | yok | taşımaz, uyarı verir |

Program her sapmayı **uyarı** olarak raporlar.

## UDF → Word

UDF'nin biçim modeli Word'ünkinin alt kümesidir; Editör'ün tanıdığı her öznitelik Word'de
karşılığıyla kurulur. Editör'ün davranışı sınıf sabitlerinden ve çizim motorunda yapılan
deneylerden öğrenilmiştir:

| Editör'de | Word'de |
|---|---|
| Alt düzey liste görünümü `NumberType`'a değil paragrafın kendi `SecListTypeLevelN` özniteliğine, yoksa varsayılana (a. / i. / (1) / (a) / (i)) bağlıdır | her `ListId` için `numbering.xml` tanımı, düzey biçimleri Editör'ün çizdiği gibi |
| `NumberSetted=n` listeyi n+1'den başlatır | `startOverride` |
| Harfli liste Türk alfabesiyle sayılır (…c, ç, d…, sonra aa, bb) | 4+ kalemli harfli listede etiketler sabit metin (Word İngiliz alfabesiyle sayar) |
| `pageNumber-spec="BSP32_n"` bir bit kümesidir: 32 ortala, 64 sağa, 2048 "sayfa/toplam" | altbilgide PAGE / NUMPAGES alanları, aynı hiza |
| `bgImage` (antet) kenar payları düşülerek her sayfaya gerilir | üstbilgiye bağlı, metnin arkasında görsel |
| satır düzeyinde `columnSpans` | ortak ızgara + `gridSpan` |
| `docx_udf.py`'nin taklitleri (zemin bandı, iç içe tabloyla dikey birleştirme) | gerçek hücre dolgusu ve `vMerge` olarak geri açılır |
| `space`, `tab`, `field`, `barcode`, tanınmayan öğeler | metin korunur, karşılığı olmayan için uyarı |

`udf_docx_testi.py` her DOCX senaryosunu UDF → DOCX → UDF turundan geçirip iki UDF'yi karşılaştırır;
ayrıca Editör'e özgü yapıları elle kurulmuş UDF'lerle sınar.

## Kaynaktan çalıştırma

Zorunlu bağımlılık yoktur (Python 3.9+):

```
python3 uygulama.py                  # pencere
python3 uygulama.py --sinama         # öz sınama
python3 docx_udf.py belge.docx -o belge.udf          # Word → UDF komut satırı
python3 md_udf.py belge.md -o belge.udf              # Markdown → UDF komut satırı
python3 udf_docx.py belge.udf -o belge.docx          # ters yön (UDF → Word)
python3 udf_dogrula.py belge.docx belge.udf          # doğrulama (sıra yönü belirler)
python3 udf_dogrula.py belge.md belge.udf            # Markdown doğrulama
python3 udf_dogrula.py belge.udf belge.docx
python3 udf_onizle.py belge.udf -o onizleme          # Editör motoruyla sayfa PNG'leri
python3 udf_senaryo_testi.py --editor                # Word → UDF senaryo testleri
python3 md_udf_testi.py                              # Markdown → UDF senaryo testleri
python3 udf_docx_testi.py                            # UDF → Word: gidiş dönüş + Editör örnekleri
```

İsteğe bağlı paketler: `tkinterdnd2` (sürükle-bırak), `pillow` (JPEG çevirisi).

## Dosyalar

| Dosya | İş |
|---|---|
| `uygulama.py` | pencere, çeviri akışı, rapor, öz sınama |
| `docx_udf.py` | DOCX → UDF dönüştürücü (çekirdek) |
| `md_udf.py` | Markdown (.md) → UDF dönüştürücü (çekirdek) |
| `udf_docx.py` | UDF → DOCX dönüştürücü (çekirdek) |
| `ofis_docx.py` | .doc/.rtf/.odt/.pages → .docx: kurulu Word (Windows'ta COM, Mac'te AppleScript) ya da Pages ile |
| `udf_dogrula.py` | doğrulama: ofset zinciri, paket bütünlüğü, tablo tutarlılığı, kelime kapsaması |
| `udf_onizle.py`, `udf_onizle/` | kurulu UYAP Editör'ün motoruyla, pencere açmadan sayfa çizimi |
| `udf_senaryo_testi.py` | Word → UDF senaryo testleri (DOCX'leri kendisi üretir) |
| `md_udf_testi.py` | Markdown → UDF senaryo testleri: başlıklar, biçimler, listeler, tablolar |
| `udf_docx_testi.py` | UDF → Word testleri: gidiş dönüş karşılaştırması ve elle kurulan Editör örnekleri |

## Koşullar

Ücretsiz kullanılabilir ve dağıtılabilir; satılamaz. Ayrıntı: [LICENSE](LICENSE).

Bu program bağımsız bir yardımcı araçtır; UYAP veya Adalet Bakanlığı ile bağlantısı yoktur.
Ürettiği belgeyi kullanmadan önce UYAP Doküman Editörü'nde açıp denetlemek kullanıcının
sorumluluğundadır.

Av. Arb. Mevlana İbrahim Asım Bilir · av.ibrahimbilir@gmail.com
