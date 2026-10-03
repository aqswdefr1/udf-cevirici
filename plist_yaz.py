#!/usr/bin/env python3
"""Derlenen .app'in Info.plist'ine sürüm, telif ve belge türü bilgisini yazar.

PyInstaller bunları komut satırından kabul etmediği için derleme sonrası
eklenir; macOS'ta Finder > Bilgi Al ve Hakkında penceresinde görünür. Belge
türü kaydı sayesinde .docx ve .udf dosyaları uygulama simgesine sürüklenebilir.
Info.plist değişince paketin imzası bozulur; derle-macos.sh ardından yeniden
(ad-hoc) imzalar.
"""
import plistlib
import sys

from uygulama import SURUM, TELIF, UYGULAMA_ADI

yol = (sys.argv[1] if len(sys.argv) > 1
       else "dist/UDF Cevirici.app/Contents/Info.plist")
with open(yol, "rb") as f:
    p = plistlib.load(f)
p["CFBundleShortVersionString"] = SURUM
p["CFBundleVersion"] = SURUM
p["CFBundleDisplayName"] = UYGULAMA_ADI
p["NSHumanReadableCopyright"] = TELIF
p["CFBundleDocumentTypes"] = [{
    "CFBundleTypeName": "Word belgesi",
    "CFBundleTypeRole": "Viewer",
    "LSHandlerRank": "Alternate",
    "LSItemContentTypes": ["org.openxmlformats.wordprocessingml.document"],
    "CFBundleTypeExtensions": ["docx"],
}, {
    "CFBundleTypeName": "Eski Word, RTF, ODT ve Pages belgesi",
    "CFBundleTypeRole": "Viewer",
    "LSHandlerRank": "Alternate",
    "CFBundleTypeExtensions": ["doc", "rtf", "odt", "pages"],
}, {
    "CFBundleTypeName": "UYAP belgesi",
    "CFBundleTypeRole": "Viewer",
    "LSHandlerRank": "Alternate",                   # UYAP Editör'ün varsayılan açıcı olmasına dokunmaz
    "CFBundleTypeExtensions": ["udf"],
}]
# .doc/.rtf/.odt/.pages belgesini Word'e ya da Pages'e Word biçimine çevirtmek için (AppleScript)
p["NSAppleEventsUsageDescription"] = ("Eski Word, RTF ve Pages belgelerini UDF'ye çevirebilmek için belgeyi "
                                      "Microsoft Word'de ya da Pages'te açıp Word biçiminde kaydettirir.")
with open(yol, "wb") as f:
    plistlib.dump(p, f)
print(f"Info.plist güncellendi: sürüm {SURUM} · {TELIF}")
