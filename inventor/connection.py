"""Autodesk Inventor.Application baglanti yonetimi.

Inventor zaten calisiyorsa ona baglanir, degilse yeni bir instance baslatir.
gencache.EnsureDispatch kullanildigi icin donen nesne early-bound'dur ve
win32com.client.constants modulu uzerinden Inventor enum sabitlerine erisim
saglar (kPartDocumentObject, kJoinOperation, vb.).
"""

import pythoncom
import win32com.client
import win32com.client.gencache

_app = None


def get_app():
    """Calisan Inventor.Application nesnesini dondurur, yoksa baslatir.

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

    # Inventor.Application COM sunucusu single-use'tir: EnsureDispatch ile
    # ProgID uzerinden istek yapildiginda, zaten calisan bir instance varsa
    # ona baglanir; yoksa yenisini baslatir. Pointer'i thread'ler arasi
    # PAYLASMIYORUZ (her cagrida taze aliniyor) - onceki hata buradan
    # kaynaklaniyordu.
    app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
    app.Visible = True
    _app = app
    return app


# Modul ilk import edildiginde Inventor'a baglanip win32com.client.constants
# modulunu Inventor enum sabitleriyle (kPartDocumentObject, kJoinOperation, ...)
# doldurur. Diger modullerdeki ust-seviye 'from win32com.client import constants'
# kullanimlari bu yuzden 'inventor.connection' import edildikten sonra calismalidir.
get_app()
