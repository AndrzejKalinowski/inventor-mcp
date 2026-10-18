"""Autodesk Inventor.Application baglanti yonetimi.

Sunucunun baslamasi Inventor'u BASLATMAZ:
- Import sirasinda sadece Inventor tip kutuphanesi (gencache.EnsureModule) yuklenir.
  Bu, win32com.client.constants modulunu Inventor enum sabitleriyle
  (kPartDocumentObject, kJoinOperation, vb.) doldurur; uygulamaya baglanmaz.
- Uygulamaya ilk tool cagrisinda baglanilir ve varsayilan olarak sadece zaten
  calisan bir Inventor'a (GetActiveObject) baglanilir. Calismiyorsa acik bir hata
  dondurulur. INVENTOR_MCP_AUTOLAUNCH=1 ile eski davranis (gerekirse baslat) secilebilir.
"""

import os
import winreg

import pythoncom
import pywintypes
import win32com.client
import win32com.client.gencache

PROG_ID = "Inventor.Application"

# "Autodesk Inventor Object Library" tip kutuphanesi. GUID'in coclass'i (_Application),
# Inventor.Application ProgID'sinin CLSID'i ile ayni (Inventor 2026 kayit defterinde dogrulandi).
TYPELIB_GUID = "{D98A091D-3A0F-4C3E-B36E-61F62068D488}"

AUTOLAUNCH_ENV = "INVENTOR_MCP_AUTOLAUNCH"

_app = None


def _autolaunch_enabled() -> bool:
    return os.environ.get(AUTOLAUNCH_ENV, "").strip().lower() in ("1", "true", "yes")


def _load_type_library():
    """Inventor tip kutuphanesini uygulamayi baslatmadan yukler (enum sabitleri icin).

    Surum, kayit defterindeki HKCR\\TypeLib\\{GUID} altindaki en yuksek surumden alinir.
    """
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "TypeLib\\" + TYPELIB_GUID) as key:
            versions = [winreg.EnumKey(key, i) for i in range(winreg.QueryInfoKey(key)[0])]
    except OSError:
        raise RuntimeError(
            f"Inventor tip kutuphanesi {TYPELIB_GUID} kayit defterinde bulunamadi - Inventor kurulu mu?"
        ) from None
    major, minor = max(tuple(int(p, 16) for p in v.split(".")) for v in versions)
    return win32com.client.gencache.EnsureModule(TYPELIB_GUID, 0, major, minor)


def get_app():
    """Calisan Inventor.Application nesnesini dondurur.

    Inventor calismiyorsa LookupError firlatir (inventor_tool bunu okunabilir hataya
    cevirir); INVENTOR_MCP_AUTOLAUNCH=1 ise yeni bir instance baslatir.

    FastMCP her tool cagrisini farkli bir thread'de calistirabilir. COM
    nesneleri apartment-threaded oldugu icin, bir thread'de olusturulan
    dispatch pointer'ini baska bir thread'den dogrudan kullanmak
    "No running instance" / com_error gibi rastgele hatalara yol acar.
    Bu yuzden her cagrida CALISTIRAN THREAD icin once CoInitialize
    yapiyoruz, sonra o thread'e ozel taze bir proxy aliyoruz - global
    pointer'i thread'ler arasi paylasmiyoruz.
    """
    global _app

    try:
        pythoncom.CoInitialize()
    except pythoncom.com_error:
        pass  # zaten bu thread'de initialize edilmis (S_FALSE)

    try:
        running = win32com.client.GetActiveObject(PROG_ID)
    except pywintypes.com_error:
        running = None

    if running is not None:
        # Early-bound proxy (CastTo ve constants icin) - EnsureDispatch mevcut nesneyi sarar.
        app = win32com.client.gencache.EnsureDispatch(running)
    elif _autolaunch_enabled():
        # Inventor.Application COM sunucusu single-use'tir: ProgID uzerinden
        # istek yapildiginda calisan instance yoksa yenisini baslatir.
        app = win32com.client.gencache.EnsureDispatch(PROG_ID)
        app.Visible = True
    else:
        raise LookupError(
            "Inventor is not running — start Inventor and try again "
            f"(or set {AUTOLAUNCH_ENV}=1 to let the server launch it)."
        )
    _app = app
    return app


# Diger modullerdeki ust-seviye 'from win32com.client import constants' kullanimlari
# (enum sozlukleri) bu modul import edildikten sonra calismalidir. Uygulamaya baglanmaz.
_load_type_library()
