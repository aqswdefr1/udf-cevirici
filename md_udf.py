#!/usr/bin/env python3
"""Markdown (.md) -> UDF (UYAP Doküman Editörü) dönüştürücü.

Markdown metinlerini veya dosyalarını, biçimlendirmelerini (başlıklar, paragraflar,
kalın/italik/altı çizili/üstü çizili/vurgu, listeler, tablolar, alıntılar, kod blokları,
görseller, dipnotlar, sayfa sonları vb.) koruyarak UYAP Doküman Editörü'nün yerel
UDF biçimine dönüştürür.

Özellikler:
  - Başlıklar: # .. ###### (H1-H6) punto ve boşluk ayarlarıyla; mahkeme başlıkları (.. MAHKEMESİNE, T.C.)
    kendiliğinden ortalanır.
  - Paragraflar: İki yana yaslama (varsayılan) veya sola/ortaya/sağa hizalama; <center>, <p align="..">,
    <div align=".."> etiketleri desteklenir.
  - Karakter Biçimleri: **kalın**, *italik*, ***kalın italik***, ~~üstü çizili~~, ==vurgu==, <u>altı çizili</u>,
    ^üst simge^, ~alt simge~, `kod`, <span style="..">, <font color=".."> vb.
  - Listeler: Madde işaretli (- , * , +) ve numaralı (1. , a. , i. , (1) , a) vb.) yerel UDF liste
    öznitelikleriyle (BulletType, NumberType, ListId, ListLevel, NumberSetted).
  - Tablolar: Markdown pipe tabloları (| col | col |), sütun hizalamaları (:---, :---:, ---:),
    başlık satırı vurgusu, otomatik sütun genişlik oranları (columnSpans).
  - Alıntılar: > Alıntı metinleri sol/sağ girinti ve italik biçimle.
  - Kod Blokları: ``` .. ``` Courier New yazı tipi, girinti ve satır aralığıyla.
  - Görseller: ![alt](yol) yerel dosyalar veya data:image URI'leri (PNG doğrudan, JPEG/BMP/GIF PNG'ye
    çevrilerek) sayfaya sığdırılarak gömülür.
  - Dipnotlar: [^1] metin içi üst simge ve belge sonu numaralı dipnot bloğu.
  - Sayfa Sonu: ---sayfa sonu---, <page-break>, \\pagebreak vb.
  - Bölücüler: ---, *** yatay çizgi satırı (rule_para).
  - Güvenlik: UTF-16 ofset zinciri denetimi, BMP dışı (emoji vb.) karakterlerin Editör uyumlu □ simgesine
    dönüştürülmesi, CDATA kaçışları.

Kullanım:
    python3 md_udf.py girdi.md -o cikti.udf
"""

import argparse
import base64
import io
import os
import re
import struct
import sys
import zipfile
from xml.etree import ElementTree as ET

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
if hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from docx_udf import (
    build, check, para_xml, _empty_para, attrs, _java_color, _xattr,
    tabset, u16, to_png, DonusumHatasi, rule_para,
    DEFAULT_FONT, HIGHLIGHT, SUBSCRIPT, IMG_CHAR,
    BULLET_ARROW, BULLET_DIAMOND, BULLET_RECT, BULLET_RECT_D, BULLET_TRIANGLE,
    _BAD, _NONBMP_MAP
)

# Mahkeme / Resmi Kurum başlıkları: kendiliğinden ortalanacak kalıplar
MAHKEME_EKLERI = (
    'MAHKEMESİNE', 'MAHKEMESİ BAŞKANLIĞINA', 'HAKİMLİĞİNE', 'SAVCILIĞINA',
    'BAŞSAVCILIĞINA', 'BAŞKANLIĞINA', 'DAİRESİNE', 'MÜDÜRLÜĞÜNE', 'KOMİSYONUNA',
    'İCRA DAİRESİNE', 'KURULUNA', 'BAKANLIĞINA', 'NOTERLİĞİNE'
)

ROMAN_VALS = {
    'i': 1, 'ii': 2, 'iii': 3, 'iv': 4, 'v': 5, 'vi': 6, 'vii': 7, 'viii': 8, 'ix': 9, 'x': 10,
    'xi': 11, 'xii': 12, 'xiii': 13, 'xiv': 14, 'xv': 15, 'xvi': 16, 'xvii': 17, 'xviii': 18, 'xix': 19, 'xx': 20
}

LIST_PAT = re.compile(
    r'^(\s*)(?:([-*+])|(?:\(([0-9a-zA-ZçÇğĞıİöÖşŞüÜ]+)\)([\.\)-]?)|(\d{1,4}|[a-zA-ZçÇğĞıİöÖşŞüÜ]|[ivxlcdm]+|[IVXLCDM]+)([\.\)-])))\s+(.*)$'
)


def _match_list_item(line):
    m = LIST_PAT.match(line)
    if not m:
        return None
    indent = m.group(1)
    if m.group(2):
        return {
            'indent': indent,
            'is_bullet': True,
            'tag': m.group(2),
            'sep': '',
            'is_paren': False,
            'text': m.group(7)
        }
    if m.group(3) is not None:
        tag = m.group(3)
        if not (tag.isdigit() or len(tag) == 1 or tag.lower() in ROMAN_VALS):
            return None
        return {
            'indent': indent,
            'is_bullet': False,
            'tag': tag,
            'sep': m.group(4) or '',
            'is_paren': True,
            'text': m.group(7)
        }
    tag = m.group(5)
    if not tag.isdigit() and len(tag) > 1 and tag.lower() not in ROMAN_VALS:
        return None
    return {
        'indent': indent,
        'is_bullet': False,
        'tag': tag,
        'sep': m.group(6) or '',
        'is_paren': False,
        'text': m.group(7)
    }


def _is_yaml_frontmatter(lines):
    if not lines:
        return False
    has_kv = False
    for line in lines:
        s = line.strip()
        if not s or s.startswith('#'):
            continue
        if re.match(r'^[a-zA-Z0-9_\-]+\s*:\s*.*$', s):
            has_kv = True
        else:
            return False
    return has_kv

# Renk eşlemeleri (CSS adları -> Hex)
CSS_RENKLER = {
    'black': '000000', 'white': 'FFFFFF', 'red': 'FF0000', 'green': '008000',
    'blue': '0000FF', 'yellow': 'FFFF00', 'cyan': '00FFFF', 'magenta': 'FF00FF',
    'gray': '808080', 'grey': '808080', 'lightgray': 'C0C0C0', 'darkgray': '404040',
    'navy': '000080', 'purple': '800080', 'orange': 'FFA500', 'brown': 'A52A2A',
    'maroon': '800000', 'olive': '808000', 'teal': '008080'
}


def hex_renk(val):
    """CSS renk değerini 6 haneli HEX koda çevirir veya None döner."""
    if not val:
        return None
    val = val.strip().lower()
    if val in CSS_RENKLER:
        return CSS_RENKLER[val]
    if val.startswith('#'):
        h = val[1:]
        if len(h) == 3:
            return ''.join(c * 2 for c in h).upper()
        if len(h) == 6:
            return h.upper()
    m = re.match(r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', val)
    if m:
        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f'{r:02X}{g:02X}{b:02X}'
    return None


def punto_boyut(val):
    """CSS uzunluk değerini (12pt, 16px vb.) punto (pt) sayısına çevirir."""
    if not val:
        return None
    m = re.match(r'([\d.]+)\s*(pt|px|em|rem)?', val.strip().lower())
    if not m:
        return None
    sayi = float(m.group(1))
    birim = m.group(2)
    if birim == 'px':
        return sayi * 0.75
    if birim in ('em', 'rem'):
        return sayi * 12.0
    return sayi


def css_stilleri_ayikla(style_str):
    """'color: red; font-size: 14pt; background-color: yellow' -> fmt dict."""
    fmt = {}
    if not style_str:
        return fmt
    kurallar = [s.strip() for s in style_str.split(';') if ':' in s]
    for k in kurallar:
        ad, deg = k.split(':', 1)
        ad, deg = ad.strip().lower(), deg.strip()
        if ad == 'color':
            c = hex_renk(deg)
            if c:
                fmt['color'] = c
        elif ad in ('background', 'background-color'):
            c = hex_renk(deg)
            if c:
                fmt['bg'] = c
        elif ad == 'font-size':
            sz = punto_boyut(deg)
            if sz:
                fmt['size'] = sz
        elif ad == 'font-family':
            fam = deg.strip('\'"').split(',')[0].strip()
            if fam:
                fmt['font'] = fam
        elif ad == 'font-weight':
            if deg in ('bold', 'bolder', '700', '800', '900'):
                fmt['bold'] = True
        elif ad == 'font-style':
            if deg in ('italic', 'oblique'):
                fmt['italic'] = True
        elif ad == 'text-decoration':
            if 'underline' in deg:
                fmt['underline'] = True
            if 'line-through' in deg:
                fmt['strike'] = True
    return fmt


# --------------------------------------------------------------------------
# Karakter / Satır içi Biçimlendirme Ayrıştırıcısı
# --------------------------------------------------------------------------

# Kaçış karakterleri için yer tutucular
_ESCAPES = {
    r'\\': '\x01_ESC_BSLASH_\x01',
    r'\*': '\x01_ESC_STAR_\x01',
    r'\_': '\x01_ESC_UNDER_\x01',
    r'\~': '\x01_ESC_TILDE_\x01',
    r'\^': '\x01_ESC_CARET_\x01',
    r'\=': '\x01_ESC_EQUAL_\x01',
    r'\`': '\x01_ESC_BACKTICK_\x01',
    r'\[': '\x01_ESC_LBRACKET_\x01',
    r'\]': '\x01_ESC_RBRACKET_\x01',
    r'\<': '\x01_ESC_LT_\x01',
    r'\>': '\x01_ESC_GT_\x01',
    r'\|': '\x01_ESC_PIPE_\x01',
    r'\#': '\x01_ESC_HASH_\x01',
    r'\!': '\x01_ESC_EXCL_\x01',
}

_UNESCAPES = {
    '\x01_ESC_BSLASH_\x01': '\\',
    '\x01_ESC_STAR_\x01': '*',
    '\x01_ESC_UNDER_\x01': '_',
    '\x01_ESC_TILDE_\x01': '~',
    '\x01_ESC_CARET_\x01': '^',
    '\x01_ESC_EQUAL_\x01': '=',
    '\x01_ESC_BACKTICK_\x01': '`',
    '\x01_ESC_LBRACKET_\x01': '[',
    '\x01_ESC_RBRACKET_\x01': ']',
    '\x01_ESC_LT_\x01': '<',
    '\x01_ESC_GT_\x01': '>',
    '\x01_ESC_PIPE_\x01': '|',
    '\x01_ESC_HASH_\x01': '#',
    '\x01_ESC_EXCL_\x01': '!',
}


def _kacis_maskele(metin):
    for k, v in _ESCAPES.items():
        metin = metin.replace(k, v)
    return metin


def _kacis_coz(metin):
    for k, v in _UNESCAPES.items():
        metin = metin.replace(k, v)
    return metin


def gorsel_yukle(url_veya_yol, base_dir, warn):
    """Görsel dosyasını veya base64 data URI'sini yükleyip PNG formatında döndürür."""
    data = None
    url = url_veya_yol.strip()
    if url.startswith('data:image/'):
        # data:[<mediatype>][;base64],<data>
        try:
            virgul = url.find(',')
            if virgul != -1:
                b64 = url[virgul + 1:]
                data = base64.b64decode(b64)
        except Exception as ex:
            warn.add(f'data URI görseli okunamadı: {ex}')
            return None
    else:
        # Yerel dosya yolu
        yol = url if os.path.isabs(url) else os.path.join(base_dir, url)
        if not os.path.exists(yol):
            warn.add(f'görsel dosyası bulunamadı: {url}')
            return None
        try:
            with open(yol, 'rb') as f:
                data = f.read()
        except OSError as ex:
            warn.add(f'görsel dosyası okunamadı: {ex}')
            return None

    if not data:
        return None

    png_data = to_png(data, url, warn)
    if not png_data or len(png_data) < 24:
        return None

    try:
        px = struct.unpack('>II', png_data[16:24])
        w_px, h_px = px[0], px[1]
        w_pt, h_pt = w_px * 0.75, h_px * 0.75
    except Exception:
        w_pt, h_pt = 200.0, 150.0

    # Sayfa genişliğine sığdır (yaklaşık 450 pt)
    max_w = 450.0
    if w_pt > max_w:
        h_pt = h_pt * max_w / w_pt
        w_pt = max_w
    max_h = 650.0
    if h_pt > max_h:
        w_pt = w_pt * max_h / h_pt
        h_pt = max_h

    b64_str = base64.encodebytes(png_data).decode('ascii')
    return {'data': b64_str, 'w': w_pt, 'h': h_pt}


def parse_inlines(metin, base_fmt=None, base_dir='.', warn=None, dipnotlar=None):
    """Markdown satır içi metnini run listesine [[metin, fmt], ...] dönüştürür."""
    if base_fmt is None:
        base_fmt = {}
    if warn is None:
        warn = set()
    if dipnotlar is None:
        dipnotlar = {}

    if not metin:
        return []

    # 1. Ters eğik çizgi kaçışlarını maskele
    metin = _kacis_maskele(metin)

    runs = []

    # Öncelikli regex kalıpları
    # Kod bloğu: `kod`
    # Görsel: ![alt](url)
    # Bağlantı: [metin](url)
    # Dipnot referansı: [^id]
    # Kalın + İtalik: ***metin*** veya ___metin___
    # Kalın: **metin** veya __metin__
    # İtalik: *metin* veya _metin_ (sözcük sınırında)
    # Üstü Çizili: ~~metin~~
    # Vurgu: ==metin==
    # Üst Simge: ^metin^
    # Alt Simge: ~metin~
    # HTML etiketleri: <span ..>..</span>, <font ..>..</font>, <b>..</b>, <i>..</i>, <u>..</u>,
    #                  <s>..</s>, <del>..</del>, <sup>..</sup>, <sub>..</sub>, <mark>..</mark>,
    #                  <img ..>, <br>

    token_pat = re.compile(
        r'(?P<code>`[^`]+`)'
        r'|(?P<img>!\[(?P<img_alt>[^\]]*)\]\((?P<img_url>[^)]+)\))'
        r'|(?P<link>\[(?P<link_text>[^\]]+)\]\((?P<link_url>[^)]+)\))'
        r'|(?P<fn>\[\^(?P<fn_id>[a-zA-Z0-9_\-]+)\])'
        r'|(?P<bi>\*\*\*(?P<bi_txt>.+?)\*\*\*|___(?P<bi_txt2>.+?)___)'
        r'|(?P<bold>\*\*(?P<bold_txt>.+?)\*\*|__(?P<bold_txt2>.+?)__)'
        r'|(?P<italic>\*(?P<it_txt>[^\*]+?)\*|(?<!\w)_(?P<it_txt2>[^_]+?)_(?!\w))'
        r'|(?P<strike>~~(?P<strike_txt>.+?)~~)'
        r'|(?P<mark>==(?P<mark_txt>.+?)==)'
        r'|(?P<sup>\^(?P<sup_txt>[^\^]+?)\^)'
        r'|(?P<sub>~(?P<sub_txt>[^~]+?)~)'
        r'|(?P<html_span><span\s+(?P<span_attrs>[^>]+)>(?P<span_txt>.*?)</span>)'
        r'|(?P<html_font><font\s+(?P<font_attrs>[^>]+)>(?P<font_txt>.*?)</font>)'
        r'|(?P<html_b><(?:b|strong)>(?P<b_txt>.*?)</(?:b|strong)>)'
        r'|(?P<html_i><(?:i|em)>(?P<i_txt>.*?)</(?:i|em)>)'
        r'|(?P<html_u><(?:u|ins)>(?P<u_txt>.*?)</(?:u|ins)>)'
        r'|(?P<html_s><(?:s|del|strike)>(?P<s_txt>.*?)</(?:s|del|strike)>)'
        r'|(?P<html_sup><sup>(?P<hsup_txt>.*?)</sup>)'
        r'|(?P<html_sub><sub>(?P<hsub_txt>.*?)</sub>)'
        r'|(?P<html_mark><mark>(?P<hmark_txt>.*?)</mark>)'
        r'|(?P<html_code><code\b[^>]*>(?P<hcode_txt>.*?)</code>)'
        r'|(?P<html_img><img\s+(?P<himg_attrs>[^>]+)>)'
        r'|(?P<html_tab><tab\s*/?>)'
        r'|(?P<html_br><br\s*/?>)',
        re.DOTALL | re.IGNORECASE
    )

    pos = 0
    for m in token_pat.finditer(metin):
        start, end = m.span()
        if start > pos:
            duz = metin[pos:start]
            runs.append([_kacis_coz(duz), dict(base_fmt)])

        g = m.groupdict()
        if g.get('code'):
            kod_metin = m.group('code')[1:-1]
            fmt = dict(base_fmt, font='Courier New', size=10.5)
            runs.append([_kacis_coz(kod_metin), fmt])

        elif g.get('img'):
            alt = m.group('img_alt')
            url = m.group('img_url')
            g_bilgi = gorsel_yukle(url, base_dir, warn)
            if g_bilgi:
                runs.append([IMG_CHAR, {'image': g_bilgi}])
            else:
                runs.append([f'[Görsel: {alt or url}]', dict(base_fmt)])

        elif g.get('link'):
            l_txt = m.group('link_text')
            l_url = m.group('link_url')
            l_fmt = dict(base_fmt, underline=True, color='0000EE')
            alt_runs = parse_inlines(l_txt, l_fmt, base_dir, warn, dipnotlar)
            runs.extend(alt_runs)

        elif g.get('fn'):
            fn_id = m.group('fn_id')
            if fn_id not in dipnotlar:
                dipnotlar[fn_id] = len(dipnotlar) + 1
            no_str = str(dipnotlar[fn_id])
            runs.append([no_str, dict(base_fmt, super=True)])

        elif g.get('bi'):
            txt = g.get('bi_txt') or g.get('bi_txt2')
            fmt = dict(base_fmt, bold=True, italic=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('bold'):
            txt = g.get('bold_txt') or g.get('bold_txt2')
            fmt = dict(base_fmt, bold=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('italic'):
            txt = g.get('it_txt') or g.get('it_txt2')
            fmt = dict(base_fmt, italic=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('strike'):
            txt = g.get('strike_txt')
            fmt = dict(base_fmt, strike=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('mark'):
            txt = g.get('mark_txt')
            fmt = dict(base_fmt, bg='FFFF00')
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('sup'):
            txt = g.get('sup_txt')
            fmt = dict(base_fmt, super=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('sub'):
            txt = g.get('sub_txt')
            fmt = dict(base_fmt, sub=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_span'):
            attrs_str = g.get('span_attrs') or ''
            txt = g.get('span_txt') or ''
            m_style = re.search(r'style\s*=\s*["\']([^"\']+)["\']', attrs_str, re.IGNORECASE)
            fmt = dict(base_fmt)
            if m_style:
                fmt.update(css_stilleri_ayikla(m_style.group(1)))
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_font'):
            attrs_str = g.get('font_attrs') or ''
            txt = g.get('font_txt') or ''
            fmt = dict(base_fmt)
            m_col = re.search(r'color\s*=\s*["\']([^"\']+)["\']', attrs_str, re.IGNORECASE)
            if m_col:
                c = hex_renk(m_col.group(1))
                if c:
                    fmt['color'] = c
            m_face = re.search(r'face\s*=\s*["\']([^"\']+)["\']', attrs_str, re.IGNORECASE)
            if m_face:
                fmt['font'] = m_face.group(1)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_b'):
            txt = g.get('b_txt') or ''
            fmt = dict(base_fmt, bold=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_i'):
            txt = g.get('i_txt') or ''
            fmt = dict(base_fmt, italic=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_u'):
            txt = g.get('u_txt') or ''
            fmt = dict(base_fmt, underline=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_s'):
            txt = g.get('s_txt') or ''
            fmt = dict(base_fmt, strike=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_sup'):
            txt = g.get('hsup_txt') or ''
            fmt = dict(base_fmt, super=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_sub'):
            txt = g.get('hsub_txt') or ''
            fmt = dict(base_fmt, sub=True)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_mark'):
            txt = g.get('hmark_txt') or ''
            fmt = dict(base_fmt, bg='FFFF00')
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_code'):
            txt = g.get('hcode_txt') or ''
            fmt = dict(base_fmt, font='Courier New', size=10.5)
            runs.extend(parse_inlines(txt, fmt, base_dir, warn, dipnotlar))

        elif g.get('html_img'):
            attrs_str = g.get('himg_attrs') or ''
            m_src = re.search(r'src\s*=\s*["\']([^"\']+)["\']', attrs_str, re.IGNORECASE)
            m_alt = re.search(r'alt\s*=\s*["\']([^"\']+)["\']', attrs_str, re.IGNORECASE)
            url = m_src.group(1) if m_src else ''
            alt = m_alt.group(1) if m_alt else ''
            g_bilgi = gorsel_yukle(url, base_dir, warn)
            if g_bilgi:
                runs.append([IMG_CHAR, {'image': g_bilgi}])
            else:
                runs.append([f'[Görsel: {alt or url}]', dict(base_fmt)])

        elif g.get('html_tab'):
            runs.append(['\t', dict(base_fmt)])

        elif g.get('html_br'):
            runs.append(['\n', {'_br': True}])

        pos = end

    if pos < len(metin):
        runs.append([_kacis_coz(metin[pos:]), dict(base_fmt)])

    # BMP dışı karakterleri ve XML geçersiz denetim karakterlerini temizle
    temiz_runs = []
    for t, f in runs:
        if 'image' in f:
            temiz_runs.append([t, f])
            continue
        # Alt simge Unicode çevirisi (Editör subscript="true" dilimini yukarıda çizdiği için)
        if f.get('sub') and t.translate(SUBSCRIPT) != t and all(ch.isspace() or ch.translate(SUBSCRIPT) != ch for ch in t):
            t = t.translate(SUBSCRIPT)
            f = {k: v for k, v in f.items() if k != 'sub'}
        # BMP dışı (emoji vb.)
        chars = []
        for ch in t:
            if ord(ch) > 0xFFFF:
                rep = _NONBMP_MAP.get(ch, chr(0x25A1))
                chars.append(rep)
                warn.add(f"Editör'ün açamadığı karakter değiştirildi: {ch}→{rep}")
            elif _BAD.match(ch):
                continue
            else:
                chars.append(ch)
        temiz_runs.append([''.join(chars), f])

    # Boş run'ları at, ardışık aynı biçimli olanları birleştir
    birlestirilmis = []
    for t, f in temiz_runs:
        if not t and 'image' not in f:
            continue
        if birlestirilmis and 'image' not in f and 'image' not in birlestirilmis[-1][1]:
            if birlestirilmis[-1][1] == f:
                birlestirilmis[-1][0] += t
                continue
        birlestirilmis.append([t, f])

    return birlestirilmis


# --------------------------------------------------------------------------
# Tablo Ayrıştırıcısı
# --------------------------------------------------------------------------

def tablo_satiri_bol(satir):
    """Tablo satırını pipe '|' işaretlerine göre böler; ters eğik çizgi ve kod bloklarını korur."""
    s = satir.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith(r'\|'):
        s = s[:-1]

    hucreler = []
    mevcut = []
    kodda = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '\\' and i + 1 < len(s):
            mevcut.append(s[i:i+2])
            i += 2
            continue
        if ch == '`':
            kodda = not kodda
            mevcut.append(ch)
        elif ch == '|' and not kodda:
            hucreler.append(''.join(mevcut).strip())
            mevcut = []
        else:
            mevcut.append(ch)
        i += 1
    hucreler.append(''.join(mevcut).strip())
    return hucreler


def tablo_ayirici_mi(satir):
    """Satırın markdown tablo ayırıcı satırı (|:---|---:|) olup olmadığını sınar."""
    s = satir.strip()
    if not ('-' in s and '|' in s):
        return False
    hucreler = tablo_satiri_bol(s)
    if not hucreler:
        return False
    for h in hucreler:
        c = h.strip()
        if not c or not re.match(r'^:?-+:?$', c):
            return False
    return True


def tablo_hizalamalari(ayirici_satir):
    """Ayırıcı satırdan her sütunun hizalama kodunu (0: sol, 1: orta, 2: sağ) döndürür."""
    hucreler = tablo_satiri_bol(ayirici_satir)
    hizalar = []
    for h in hucreler:
        c = h.strip()
        if c.startswith(':') and c.endswith(':'):
            hizalar.append(1)  # orta
        elif c.endswith(':'):
            hizalar.append(2)  # sağ
        else:
            hizalar.append(0)  # sol
    return hizalar


# --------------------------------------------------------------------------
# Belge Blokları Ayrıştırıcısı
# --------------------------------------------------------------------------

def parse_markdown(md_metin, base_dir='.', varsayilan_hiza=3, h1_ortala=True, warn=None):
    """Markdown metnini UDF bloklarına dönüştürür.

    Döner: (blocks, info, hf, warn)
    """
    if warn is None:
        warn = set()

    satirlar = md_metin.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    blocks = []
    dipnot_metinleri = {}
    dipnot_sirasi = {}

    # 1. YAML Frontmatter denetimi
    idx = 0
    baslik_etiketi = None
    if satirlar and satirlar[0].strip() == '---':
        idx = 1
        fm_satirlari = []
        while idx < len(satirlar) and satirlar[idx].strip() not in ('---', '...'):
            fm_satirlari.append(satirlar[idx])
            idx += 1
        if idx < len(satirlar) and _is_yaml_frontmatter(fm_satirlari):
            idx += 1  # kapatıcı '---' satırını geç
            # Frontmatter anahtarlarını basitçe tara
            for fml in fm_satirlari:
                if ':' in fml:
                    k, v = fml.split(':', 1)
                    k = k.strip().lower()
                    v = v.strip().strip('"\'')
                    if k in ('title', 'baslik'):
                        baslik_etiketi = v
        else:
            idx = 0  # Gerçek frontmatter değil; yatay çizgi ve içeriği koru

    satirlar = satirlar[idx:]

    # 2. Dipnot tanımlarını önceden topla: [^1]: Açıklama
    kalan_satirlar = []
    current_fn_id = None
    for s in satirlar:
        m_fn = re.match(r'^\[\^([a-zA-Z0-9_\-]+)\]:\s*(.*)$', s.strip())
        if m_fn:
            current_fn_id = m_fn.group(1)
            dipnot_metinleri[current_fn_id] = m_fn.group(2).strip()
        elif current_fn_id and (s.startswith('    ') or s.startswith('\t')) and s.strip():
            dipnot_metinleri[current_fn_id] += ' ' + s.strip()
        else:
            current_fn_id = None
            kalan_satirlar.append(s)
    satirlar = kalan_satirlar

    # Belge başlığı frontmatter'dan geldiyse en başa ekle
    if baslik_etiketi:
        p = _empty_para()
        p.update(
            align=1,
            runs=[[baslik_etiketi, {'bold': True, 'size': 16}]],
            before=14.0, after=8.0
        )
        blocks.append(p)

    list_id_sayaci = 1

    i = 0
    n = len(satirlar)
    while i < n:
        satir = satirlar[i]
        satir_str = satir.strip()

        # Boş satır
        if not satir_str:
            i += 1
            continue

        # --- A. Sayfa Sonu ---
        if (satir_str.lower() in ('---sayfa sonu---', '---sayfa-sonu---', '\\pagebreak', '\\newpage')
                or re.match(r'^(?:<!--\s*page[\s\-_]*break\s*-->|<page-break\s*/?>|<pagebreak\s*/?>)$', satir_str, re.IGNORECASE)):
            blocks.append({'type': 'pb'})
            i += 1
            continue

        # --- B. Yatay Çizgi / Bölücü ---
        if re.match(r'^(?:-{3,}|\*{3,}|_{3,})$', satir_str):
            blocks.append(rule_para(color='000000', text_w=450.0))
            i += 1
            continue

        # --- C. Fenced Kod Bloğu ``` veya ~~~ ---
        if satir_str.startswith('```') or satir_str.startswith('~~~'):
            fence = satir_str[:3]
            i += 1
            kod_satirlari = []
            while i < n and not satirlar[i].strip().startswith(fence):
                kod_satirlari.append(satirlar[i])
                i += 1
            if i < n:
                i += 1  # kapatıcı ```
            for kl in kod_satirlari:
                p = _empty_para()
                p.update(
                    align=0,
                    runs=[[kl if kl else ' ', {'font': 'Courier New', 'size': 10}]],
                    indent={'left': 18.0},
                    before=1.0, after=1.0, line=0.0
                )
                blocks.append(p)
            continue

        # --- D. Setext Başlıklar (Satırın altındaki === veya ---) ---
        if i + 1 < n and satirlar[i + 1].strip() and not satir_str.startswith(('#', '-', '*', '+', '>')):
            alt_satir = satirlar[i + 1].strip()
            if re.match(r'^={3,}$', alt_satir):  # H1
                p = _empty_para()
                al = 1 if h1_ortala else 0
                runs = parse_inlines(satir_str, {'bold': True, 'size': 16}, base_dir, warn, dipnot_sirasi)
                p.update(align=al, runs=runs, before=14.0, after=6.0)
                blocks.append(p)
                i += 2
                continue
            elif re.match(r'^-{3,}$', alt_satir) and not tablo_ayirici_mi(alt_satir):  # H2
                p = _empty_para()
                runs = parse_inlines(satir_str, {'bold': True, 'size': 14}, base_dir, warn, dipnot_sirasi)
                p.update(align=0, runs=runs, before=10.0, after=4.0)
                blocks.append(p)
                i += 2
                continue

        # --- E. ATX Başlıklar (# .. ######) ---
        m_h = re.match(r'^(#{1,6})\s+(.*)$', satir_str)
        if m_h:
            seviye = len(m_h.group(1))
            b_metin = m_h.group(2).strip().rstrip('#').strip()

            # Hizalama tespiti:
            # - Başlık metninde <center> etiketi var mı?
            # - Mahkeme adı mı? (.. MAHKEMESİNE vb.)
            # - H1 ortalama seçeneği açık mı?
            al = 0
            if ('<center>' in b_metin.lower() or b_metin.startswith('->')
                    or any(b_metin.upper().endswith(ek) for ek in MAHKEME_EKLERI)
                    or b_metin.upper() in ('T.C.', 'T.C')
                    or (seviye == 1 and h1_ortala)):
                al = 1

            # Temizle <center> etiketlerini
            b_metin = re.sub(r'</?center>', '', b_metin, flags=re.IGNORECASE).strip()
            if b_metin.startswith('->') and b_metin.endswith('<-'):
                b_metin = b_metin[2:-2].strip()

            boyutlar = {1: (16, 14.0, 6.0), 2: (14, 10.0, 4.0), 3: (13, 8.0, 3.0),
                        4: (12, 6.0, 2.0), 5: (12, 4.0, 2.0), 6: (12, 4.0, 2.0)}
            sz, bef, aft = boyutlar.get(seviye, (12, 4.0, 2.0))
            is_it = True if seviye in (5, 6) else False

            p = _empty_para()
            runs = parse_inlines(b_metin, {'bold': True if seviye <= 5 else False, 'italic': is_it, 'size': sz},
                                 base_dir, warn, dipnot_sirasi)
            p.update(align=al, runs=runs, before=bef, after=aft)
            blocks.append(p)
            i += 1
            continue

        # --- F. Tablolar (| Başlık | .. |) ---
        if '|' in satir and i + 1 < n and tablo_ayirici_mi(satirlar[i + 1]):
            baslik_satiri = satir
            ayirici_satir = satirlar[i + 1]
            hizalar = tablo_hizalamalari(ayirici_satir)
            ncol = len(hizalar)

            b_hucreler = tablo_satiri_bol(baslik_satiri)
            while len(b_hucreler) < ncol:
                b_hucreler.append('')
            b_hucreler = b_hucreler[:ncol]

            veri_satirlari = []
            i += 2
            while i < n and satirlar[i].strip() and '|' in satirlar[i]:
                vh = tablo_satiri_bol(satirlar[i])
                while len(vh) < ncol:
                    vh.append('')
                veri_satirlari.append(vh[:ncol])
                i += 1

            # Sütun oranları (columnSpans) hesabı
            agirliklar = []
            tum_satirlar = [b_hucreler] + veri_satirlari
            for c in range(ncol):
                max_u = max(len(s[c]) for s in tum_satirlar) if tum_satirlar else 4
                agirliklar.append(max(max_u, 4))
            top_ag = sum(agirliklar) or 1
            spans = [max(1, int(round(w / top_ag * 100 * ncol))) for w in agirliklar]
            if spans:
                spans[-1] += 100 * ncol - sum(spans)

            def _hucre_kur(hucre_metin, col_hiza, in_header):
                al = col_hiza
                m_str = hucre_metin.strip()
                if '<center>' in m_str.lower():
                    al = 1
                elif re.search(r'<(?:p|div)\s+align=["\']right["\']', m_str, re.IGNORECASE):
                    al = 2
                elif re.search(r'<(?:p|div)\s+align=["\']center["\']', m_str, re.IGNORECASE):
                    al = 1
                elif re.search(r'<(?:p|div)\s+align=["\']left["\']', m_str, re.IGNORECASE):
                    al = 0
                elif re.search(r'<(?:p|div)\s+align=["\']justify["\']', m_str, re.IGNORECASE):
                    al = 3

                m_temiz = re.sub(r'</?(?:center|p|div)(?:\s+[^>]*)?>', '', m_str, flags=re.IGNORECASE).strip()
                alt_parcalar = re.split(r'<br\s*/?>', m_temiz, flags=re.IGNORECASE)
                paras = []
                bef = 3.0 if in_header else 2.0
                aft = 3.0 if in_header else 2.0
                base_f = {'bold': True} if in_header else {}
                for sp in alt_parcalar:
                    sp_str = sp.strip()
                    r_runs = parse_inlines(sp_str, base_f, base_dir, warn, dipnot_sirasi)
                    cp = _empty_para(in_cell=True)
                    cp.update(align=al, runs=r_runs if r_runs else [['', {}]], before=bef, after=aft)
                    paras.append(cp)
                return {'paras': paras or [_empty_para(in_cell=True)], 'span': 1, 'vmerge': None, 'valign': 'center', 'fill': None}

            # Başlık satırı hücreleri
            tbl_rows = []
            h_cells = [_hucre_kur(h_metin, hizalar[col_idx], True) for col_idx, h_metin in enumerate(b_hucreler)]
            tbl_rows.append({'cells': h_cells, 'min_h': 0})

            # Veri satırları hücreleri
            for v_satir in veri_satirlari:
                v_cells = [_hucre_kur(v_metin, hizalar[col_idx], False) for col_idx, v_metin in enumerate(v_satir)]
                tbl_rows.append({'cells': v_cells, 'min_h': 0})

            blocks.append({
                'type': 'tbl',
                'ncol': ncol,
                'spans': spans,
                'rows': tbl_rows,
                'border': 'borderCell'
            })
            continue

        # --- G. Alıntılar (> Alıntı..) ---
        if satir_str.startswith('>'):
            q_satirlari = []
            seviye = 1
            while i < n and satirlar[i].strip().startswith('>'):
                m_q = re.match(r'^\s*(>+)\s*(.*)$', satirlar[i].strip())
                if m_q:
                    seviye = max(seviye, len(m_q.group(1)))
                    q_satirlari.append(m_q.group(2))
                i += 1
            q_metin = ' '.join(q_satirlari)
            p = _empty_para()
            runs = parse_inlines(q_metin, {'italic': True}, base_dir, warn, dipnot_sirasi)
            p.update(
                align=varsayilan_hiza,
                runs=runs,
                indent={'left': 36.0 * seviye, 'right': 18.0},
                before=4.0, after=4.0
            )
            blocks.append(p)
            continue

        # --- H. Listeler (- Madde, 1. Madde, a) Bent vb.) ---
        m_list = _match_list_item(satir)
        if m_list:
            list_id = list_id_sayaci
            list_id_sayaci += 1
            ilk_kalem = True
            prev_is_bullet = m_list['is_bullet']
            char_counts = {}

            while i < n:
                cur_line = satirlar[i]
                if not cur_line.strip():
                    break

                item_info = _match_list_item(cur_line)
                if not item_info:
                    # Liste devam satırı (girintili paragraf)
                    if blocks and blocks[-1]['type'] == 'p' and blocks[-1].get('list'):
                        ek_runs = parse_inlines(' ' + cur_line.strip(), {}, base_dir, warn, dipnot_sirasi)
                        blocks[-1]['runs'].extend(ek_runs)
                        i += 1
                        continue
                    else:
                        break

                girinti_str = item_info['indent']
                seviye = min(6, max(1, len(girinti_str) // 2 + 1))
                is_bullet = item_info['is_bullet']
                madde_metin = item_info['text']

                # Düzey 1'de madde işareti türü (numaralı vs simgeli) değişirse yeni list_id başlat
                if not ilk_kalem and seviye == 1 and is_bullet != prev_is_bullet:
                    list_id = list_id_sayaci
                    list_id_sayaci += 1
                    ilk_kalem = True
                    char_counts = {}

                prev_is_bullet = is_bullet
                p = _empty_para()
                runs = parse_inlines(madde_metin, {}, base_dir, warn, dipnot_sirasi)

                if is_bullet:
                    bt_map = {1: 'ELLIPSE', 2: 'RECTANGLE', 3: 'RECTANGLE_D', 4: 'ARROW'}
                    bt = bt_map.get(seviye, 'ELLIPSE')
                    sec = f' SecListTypeLevel{seviye}="BULLET_TYPE_{bt}"' if 2 <= seviye <= 6 else ''
                    list_attr = f' Bulleted="true" BulletType="BULLET_TYPE_{bt}"{sec} ListId="{list_id}" ListLevel="{seviye}"'
                else:
                    etiket = item_info['tag']
                    ayrac = item_info['sep']
                    is_paren = item_info['is_paren']

                    fam = 'NUMBER'
                    suf = 'DOT'
                    if is_paren:
                        suf = 'D_PARANTHESE'
                    elif ayrac == '.':
                        suf = 'DOT'
                    elif ayrac == ')':
                        suf = 'PARANTHESE'
                    elif ayrac == '-':
                        suf = 'TRE'

                    start_num = 1
                    if etiket.isdigit():
                        fam = 'NUMBER'
                        start_num = int(etiket)
                    elif len(etiket) > 1 and etiket.lower() in ROMAN_VALS:
                        fam = 'ROMAN_SMALL' if etiket.islower() else 'ROMAN_BIG'
                        start_num = ROMAN_VALS[etiket.lower()]
                    elif len(etiket) == 1:
                        if is_paren and etiket.lower() == 'i':
                            fam = 'ROMAN_SMALL' if etiket.islower() else 'ROMAN_BIG'
                            start_num = 1
                        elif etiket.islower():
                            fam = 'CHAR_SMALL'
                        else:
                            fam = 'CHAR_BIG'
                    elif etiket.lower() in ROMAN_VALS:
                        fam = 'ROMAN_SMALL' if etiket.islower() else 'ROMAN_BIG'
                        start_num = ROMAN_VALS[etiket.lower()]

                    if fam in ('CHAR_SMALL', 'CHAR_BIG'):
                        char_counts[seviye] = char_counts.get(seviye, 0) + 1
                        if char_counts[seviye] >= 4:
                            warn.add("harfli listede Editör Türk alfabesiyle sayar (…c, ç, d…): "
                                     "Word'deki 4. ve sonraki harfler (d, e…) Editör'de bir harf kayar")

                    typ = f'NUMBER_TYPE_{fam}_{suf}'
                    sec = f' SecListTypeLevel{seviye}="{typ}"' if 2 <= seviye <= 6 else ''
                    setted = f' NumberSetted="{start_num - 1}"' if (start_num > 1 and ilk_kalem) else ''
                    list_attr = f' Numbered="true" NumberType="{typ}"{sec}{setted} ListId="{list_id}" ListLevel="{seviye}"'

                p.update(
                    align=0,
                    runs=runs,
                    list=list_attr,
                    indent={'left': 36.0 * seviye},
                    before=1.5, after=1.5
                )
                blocks.append(p)
                ilk_kalem = False
                i += 1
            continue

        # --- I. Düz Paragraf ---
        # Ardışık boş olmayan satırları birleştir
        para_satirlari = []
        while i < n and satirlar[i].strip():
            # Araya giren başka blokları kontrol et
            s_test = satirlar[i].strip()
            if (s_test.startswith(('#', '```', '~~~', '---sayfa'))
                    or re.match(r'^(?:-{3,}|\*{3,}|_{3,})$', s_test)
                    or re.match(r'^(?:<!--\s*pagebreak\s*-->|<page-break\s*/?>)$', s_test, re.IGNORECASE)
                    or (i + 1 < n and tablo_ayirici_mi(satirlar[i + 1]))):
                break
            para_satirlari.append(satirlar[i])
            i += 1

        if not para_satirlari:
            i += 1
            continue

        # Paragraf içi satır sonları (<br> veya satır sonu 2 boşluk)
        alt_paragraflar = []
        mevcut_parca = []
        for pl in para_satirlari:
            has_break = pl.endswith('  ') or pl.strip().endswith('<br>') or pl.strip().endswith('<br/>')
            temiz = re.sub(r'<br\s*/?>\s*$', '', pl).strip()
            mevcut_parca.append(temiz)
            if has_break:
                alt_paragraflar.append(' '.join(mevcut_parca))
                mevcut_parca = []
        if mevcut_parca:
            alt_paragraflar.append(' '.join(mevcut_parca))

        for pi_idx, p_metin in enumerate(alt_paragraflar):
            if not p_metin.strip():
                continue

            # Paragraf düzeyinde hizalama ve biçim tespiti
            al = varsayilan_hiza
            p_strip = p_metin.strip()

            # Mahkeme / Resmi Kurum başlığı mı?
            if (any(p_strip.upper().endswith(ek) for ek in MAHKEME_EKLERI)
                    or p_strip.upper() in ('T.C.', 'T.C')):
                al = 1
                base_f = {'bold': True}
            elif '<center>' in p_strip.lower() or (p_strip.startswith('->') and p_strip.endswith('<-')):
                al = 1
                base_f = {}
            elif re.search(r'<(?:p|div)\s+align=["\']center["\']', p_strip, re.IGNORECASE):
                al = 1
                base_f = {}
            elif re.search(r'<(?:p|div)\s+align=["\']right["\']', p_strip, re.IGNORECASE):
                al = 2
                base_f = {}
            elif re.search(r'<(?:p|div)\s+align=["\']left["\']', p_strip, re.IGNORECASE):
                al = 0
                base_f = {}
            elif re.search(r'<(?:p|div)\s+align=["\']justify["\']', p_strip, re.IGNORECASE):
                al = 3
                base_f = {}
            else:
                base_f = {}

            # Blok etiketlerini metinden temizle
            p_temiz = re.sub(r'</?(?:center|p|div)(?:\s+[^>]*)?>', '', p_strip, flags=re.IGNORECASE).strip()
            if p_temiz.startswith('->') and p_temiz.endswith('<-'):
                p_temiz = p_temiz[2:-2].strip()

            p = _empty_para()
            runs = parse_inlines(p_temiz, base_f, base_dir, warn, dipnot_sirasi)

            # Tab durağı tespiti (DAVACI\t: Ahmet Yılmaz gibi kalıplar)
            tabs = None
            if '\t' in p_temiz:
                tabs = [(144.0, 'left', None)]

            bef = 4.0 if pi_idx == 0 else 1.0
            aft = 4.0 if pi_idx == len(alt_paragraflar) - 1 else 1.0

            p.update(
                align=al,
                runs=runs,
                tabs=tabs,
                before=bef, after=aft
            )
            blocks.append(p)

    # 3. Dipnotlar Bölümü (Varsa belge sonuna ekle)
    kalan_fn = [fn_id for fn_id in dipnot_metinleri if fn_id not in dipnot_sirasi]
    if dipnot_sirasi or kalan_fn:
        # Küçük bir ayırıcı çizgi
        blocks.append(rule_para(color='808080', text_w=150.0))
        # Sıraya göre dipnotları yaz
        sirali_dn = sorted(dipnot_sirasi.items(), key=lambda x: x[1])
        next_no = (sirali_dn[-1][1] + 1) if sirali_dn else 1
        for k_id in kalan_fn:
            sirali_dn.append((k_id, next_no))
            next_no += 1
        for fn_id, no in sirali_dn:
            dn_metin = dipnot_metinleri.get(fn_id, '')
            p = _empty_para()
            runs = [
                [f'{no}. ', {'bold': True, 'size': 10}],
            ]
            runs.extend(parse_inlines(dn_metin, {'size': 10}, base_dir, warn))
            p.update(
                align=varsayilan_hiza,
                runs=runs,
                indent={'left': 18.0},
                before=2.0, after=2.0
            )
            blocks.append(p)

    # 4. Blok normalizasyonu (tablolar arasına ve belgenin sonuna paragraf ekleme)
    fixed = []
    for blk in blocks:
        if blk['type'] == 'tbl' and fixed and fixed[-1]['type'] == 'tbl':
            fixed.append(_empty_para())
        fixed.append(blk)
    if not fixed or fixed[-1]['type'] != 'p':
        fixed.append(_empty_para())
    blocks = fixed

    # Belge sayfa bilgisi (Standart A4)
    info = {
        'page_w': 595.3,
        'page_h': 841.9,
        'margins': [70.85, 70.85, 70.85, 70.85],
        'media': 1,
        'landscape': False,
        'bg': None,
        'text_w': 453.6,
        'text_h': 700.2
    }
    hf = {'header': [], 'footer': []}

    return blocks, info, hf, warn


# --------------------------------------------------------------------------
# Çeviri Ana Fonksiyonu
# --------------------------------------------------------------------------

def cevir(kaynak, cikti, sayfa_no='otomatik', hiza='ikiye-yasla', h1_ortala=True):
    """Markdown dosyasını veya metnini UDF biçimine çevirir.

    sayfa_no: 'otomatik' (varsayılan: altbilgide sayfa no var) | 'var' | 'yok'
    hiza: 'ikiye-yasla' (3) | 'sola' (0)
    h1_ortala: True | False (H1 başlıklarını ortala)
    """
    warn = set()

    # Kaynak dosya mı yoksa doğrudan metin mi?
    if os.path.exists(kaynak):
        base_dir = os.path.dirname(os.path.abspath(kaynak))
        try:
            with open(kaynak, 'r', encoding='utf-8', errors='replace') as f:
                md_metin = f.read()
        except OSError as ex:
            raise DonusumHatasi(f'Markdown dosyası okunamadı: {ex}')
    elif '\n' not in kaynak and (kaynak.lower().endswith(('.md', '.markdown')) or os.path.sep in kaynak or '/' in kaynak):
        raise DonusumHatasi(f'dosya bulunamadı: {kaynak}')
    else:
        # Metin olarak değerlendir
        base_dir = os.getcwd()
        md_metin = kaynak

    v_hiza = 3 if hiza == 'ikiye-yasla' else 0
    blocks, info, hf, warn = parse_markdown(md_metin, base_dir, v_hiza, h1_ortala, warn)

    xml, cdata = build(blocks, info, hf, sayfa_no)
    check(xml, cdata)

    try:
        with zipfile.ZipFile(cikti, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('content.xml', xml.encode('utf-8'))
    except OSError as ex:
        raise DonusumHatasi(f'UDF dosyası yazılamadı: {ex}')

    # Görsel sayısını hesapla
    gorsel_sayisi = 0
    for b in blocks:
        if b['type'] == 'p':
            gorsel_sayisi += sum(1 for _, f in b.get('runs', []) if 'image' in f)
        elif b['type'] == 'tbl':
            for r in b.get('rows', []):
                for c in r.get('cells', []):
                    for cp in c.get('paras', []):
                        gorsel_sayisi += sum(1 for _, f in cp.get('runs', []) if 'image' in f)

    return {
        'paragraf': sum(1 for b in blocks if b['type'] == 'p'),
        'tablo': sum(1 for b in blocks if b['type'] == 'tbl'),
        'sayfa_sonu': sum(1 for b in blocks if b['type'] == 'pb'),
        'karakter': u16(cdata),
        'gorsel': gorsel_sayisi,
        'ustbilgi': bool(hf.get('header')),
        'altbilgi': bool(hf.get('footer')) or (sayfa_no != 'yok'),
        'uyarilar': sorted(warn)
    }


def main():
    ap = argparse.ArgumentParser(description='Markdown (.md) belgesini UYAP UDF biçimine çevirir.')
    ap.add_argument('md', help='Girdi Markdown (.md) belgesi')
    ap.add_argument('-o', '--output', help='Çıktı UDF (.udf) dosyası (varsayılan: <ad>.udf)')
    ap.add_argument('--sayfa-no', choices=('otomatik', 'var', 'yok'), default='otomatik',
                    help='Altbilgi sayfa numarası: otomatik (varsayılan), var, yok')
    ap.add_argument('--hiza', choices=('ikiye-yasla', 'sola'), default='ikiye-yasla',
                    help='Paragraf varsayılan hizalaması: ikiye-yasla (varsayılan), sola')
    ap.add_argument('--h1-sol', action='store_true',
                    help='H1 başlıklarını ortalamak yerine sola yasla')

    args = ap.parse_args()
    cikti = args.output or (os.path.splitext(args.md)[0] + '.udf')
    try:
        ozet = cevir(args.md, cikti, sayfa_no=args.sayfa_no, hiza=args.hiza, h1_ortala=not args.h1_sol)
    except DonusumHatasi as ex:
        sys.exit(f'HATA: {ex}')

    ekler = []
    if ozet['sayfa_sonu']:
        ekler.append(f'{ozet["sayfa_sonu"]} sayfa sonu')
    if ozet['gorsel']:
        ekler.append(f'{ozet["gorsel"]} görsel')
    if ozet['altbilgi']:
        ekler.append('altbilgi')
    ek_str = (', ' + ', '.join(ekler)) if ekler else ''

    print(f'Yazıldı: {cikti} ({ozet["paragraf"]} paragraf, {ozet["tablo"]} tablo, '
          f'{ozet["karakter"]} karakter{ek_str})')
    for u in ozet['uyarilar']:
        print('UYARI:', u)


if __name__ == '__main__':
    main()
