"""Ortak yardimcilar: dokuman/sketch/feature lookup, hata sarmalama, transient geometry."""

import functools
import threading

import pywintypes
from win32com.client import CastTo, constants

# Inventor COM nesnesi tek-threaded (STA). FastMCP eszamanli/paralel tool
# cagrilarini farkli thread'lerde calistirabilir; bu, "yanlis dokuman
# kapandi" gibi yarış (race condition) hatalarina yol acabilir (gercekten
# yasandi - bkz. inventor_close_document olayı). Bu kilit, ayni anda
# sadece BIR Inventor COM cagrisinin yapilmasini garanti eder.
_com_lock = threading.Lock()

from inventor.units import mm_to_cm

DOC_TYPE_NAMES = {
    constants.kPartDocumentObject: "part",
    constants.kAssemblyDocumentObject: "assembly",
    constants.kDrawingDocumentObject: "drawing",
    constants.kPresentationDocumentObject: "presentation",
}

DOC_TYPE_FROM_NAME = {
    "part": constants.kPartDocumentObject,
    "assembly": constants.kAssemblyDocumentObject,
    "drawing": constants.kDrawingDocumentObject,
}


def format_com_error(e: pywintypes.com_error) -> str:
    """pywintypes.com_error nesnesini okunabilir bir mesaja cevirir."""
    hresult = e.args[0]
    msg = e.args[1] if len(e.args) > 1 else ""
    excepinfo = e.args[2] if len(e.args) > 2 else None
    detail = ""
    if excepinfo:
        source = excepinfo[1] if len(excepinfo) > 1 else None
        description = excepinfo[2] if len(excepinfo) > 2 else None
        if description:
            detail = f" -- {source}: {description}" if source else f" -- {description}"
    return f"Inventor COM hatasi 0x{hresult & 0xFFFFFFFF:08X}: {msg}{detail}"


def inventor_tool(func):
    """COM ve lookup hatalarini okunabilir RuntimeError'a ceviren decorator."""

    @functools.wraps(func)
    async def wrapper(*args, **kwargs):
        with _com_lock:
            try:
                return await func(*args, **kwargs)
            except pywintypes.com_error as e:
                raise RuntimeError(format_com_error(e)) from None
            except LookupError as e:
                raise RuntimeError(str(e)) from None

    return wrapper


def doc_type_name(doc) -> str:
    return DOC_TYPE_NAMES.get(doc.DocumentType, str(doc.DocumentType))


def get_document(app, name: str | None = None):
    """Doküman adına veya tam dosya yoluna göre açık bir dokümanı bulur.

    name verilmezse aktif dokümanı döndürür.
    """
    if name is None:
        doc = app.ActiveDocument
        if doc is None:
            raise LookupError("Acik bir dokuman yok ve 'name' belirtilmedi.")
        return doc

    for d in app.Documents:
        if d.DisplayName == name or d.FullFileName == name:
            return d
    raise LookupError(f"'{name}' adli acik bir dokuman bulunamadi.")


def get_part_component_def(doc):
    """Bir Part dokumaninin ComponentDefinition'ini dondurur.

    `Documents.Add`/`Documents.Item` gibi genel donus tipli COM cagrilari erken-bagli
    generic 'Document' arayuzunu dondurur; bu arayuzde 'ComponentDefinition' yoktur.
    Bu yuzden somut 'PartDocument' arayuzune CastTo ile gecis yapilir.
    """
    if doc.DocumentType != constants.kPartDocumentObject:
        raise LookupError(
            f"Bu islem sadece Part (.ipt) dokumanlarinda calisir, '{doc.DisplayName}' bir {doc_type_name(doc)} dokumani."
        )
    part_doc = CastTo(doc, "PartDocument")
    return part_doc.ComponentDefinition


def get_assembly_component_def(doc):
    """Bir Assembly dokumaninin AssemblyComponentDefinition'ini dondurur."""
    if doc.DocumentType != constants.kAssemblyDocumentObject:
        raise LookupError(
            f"Bu islem sadece Assembly (.iam) dokumanlarinda calisir, '{doc.DisplayName}' bir {doc_type_name(doc)} dokumani."
        )
    asm_doc = CastTo(doc, "AssemblyDocument")
    return asm_doc.ComponentDefinition


def get_drawing_document(doc):
    """Bir Drawing dokumanini somut 'DrawingDocument' arayuzune cevirir."""
    if doc.DocumentType != constants.kDrawingDocumentObject:
        raise LookupError(
            f"Bu islem sadece Drawing (.idw/.dwg) dokumanlarinda calisir, '{doc.DisplayName}' bir {doc_type_name(doc)} dokumani."
        )
    return CastTo(doc, "DrawingDocument")


def get_sheet(dwg_doc, name: str | None = None):
    """Bir Drawing dokumanindaki sheet'i adina gore bulur; verilmezse aktif sheet'i dondurur."""
    if name is None:
        return dwg_doc.ActiveSheet
    try:
        return dwg_doc.Sheets.Item(name)
    except pywintypes.com_error:
        raise LookupError(f"'{name}' adli sheet bulunamadi.")


def get_sketch(doc, name: str):
    """Part dokumanindaki bir PlanarSketch'i adina gore bulur."""
    comp_def = get_part_component_def(doc)
    for sk in comp_def.Sketches:
        if sk.Name == name:
            return sk
    raise LookupError(f"'{name}' adli sketch bulunamadi.")


def get_feature(doc, name: str):
    """Part dokumanindaki bir feature'i adina gore bulur."""
    comp_def = get_part_component_def(doc)
    try:
        return comp_def.Features.Item(name)
    except pywintypes.com_error:
        raise LookupError(f"'{name}' adli feature bulunamadi.")


def point2d(app, x_mm: float, y_mm: float):
    """mm cinsinden 2D nokta olusturur (Inventor icinde cm'e cevrilir)."""
    tg = app.TransientGeometry
    return tg.CreatePoint2d(mm_to_cm(x_mm), mm_to_cm(y_mm))


def point3d(app, x_mm: float, y_mm: float, z_mm: float):
    """mm cinsinden 3D nokta olusturur (Inventor icinde cm'e cevrilir)."""
    tg = app.TransientGeometry
    return tg.CreatePoint(mm_to_cm(x_mm), mm_to_cm(y_mm), mm_to_cm(z_mm))
