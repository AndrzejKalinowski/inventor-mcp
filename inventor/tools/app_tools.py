"""Uygulama ve dokuman yonetimi araclari."""

from inventor.connection import get_app
from inventor.helpers import DOC_TYPE_FROM_NAME, doc_type_name, get_document, inventor_tool


def _doc_summary(d) -> dict:
    return {
        "name": d.DisplayName,
        "type": doc_type_name(d),
        "full_path": d.FullFileName,
        "dirty": bool(d.Dirty),
    }


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_get_status() -> dict:
        """Inventor uygulamasinin durumunu dondurur: calisiyor mu, surum, gorunurluk,
        acik dokumanlar ve aktif dokuman."""
        app = get_app()
        active_doc = app.ActiveDocument
        return {
            "running": True,
            "visible": bool(app.Visible),
            "version": app.SoftwareVersion.DisplayVersion,
            "documents": [_doc_summary(d) for d in app.Documents],
            "active_document": active_doc.DisplayName if active_doc else None,
        }

    @mcp.tool()
    @inventor_tool
    async def inventor_list_documents() -> list[dict]:
        """Acik tum Inventor dokumanlarini listeler (ad, tur, dosya yolu, kaydedilmemis degisiklik)."""
        app = get_app()
        return [_doc_summary(d) for d in app.Documents]

    @mcp.tool()
    @inventor_tool
    async def inventor_new_document(doc_type: str = "part") -> dict:
        """Yeni bir Inventor dokumani olusturur.

        doc_type: 'part' (.ipt), 'assembly' (.iam) veya 'drawing' (.idw/.dwg).
        """
        app = get_app()
        const = DOC_TYPE_FROM_NAME.get(doc_type.lower())
        if const is None:
            raise LookupError(
                f"Gecersiz doc_type: '{doc_type}'. 'part', 'assembly' veya 'drawing' olmali."
            )
        template = app.FileManager.GetTemplateFile(const)
        doc = app.Documents.Add(const, template, True)
        return _doc_summary(doc)

    @mcp.tool()
    @inventor_tool
    async def inventor_open_document(path: str) -> dict:
        """Var olan bir Inventor dosyasini (.ipt/.iam/.idw/.dwg) acar."""
        app = get_app()
        doc = app.Documents.Open(path, True)
        return _doc_summary(doc)

    @mcp.tool()
    @inventor_tool
    async def inventor_save_document(name: str | None = None, save_as_path: str | None = None) -> dict:
        """Dokumani kaydeder.

        name verilmezse aktif dokuman kaydedilir. save_as_path verilirse o yola
        farkli kaydedilir (Save As).
        """
        app = get_app()
        doc = get_document(app, name)
        if save_as_path:
            doc.SaveAs(save_as_path, False)
        else:
            doc.Save()
        return _doc_summary(doc)

    @mcp.tool()
    @inventor_tool
    async def inventor_close_document(name: str | None = None, save: bool = True) -> dict:
        """Dokumani kapatir. name verilmezse aktif dokuman kapatilir.

        save=True ise kapatmadan once kaydedilmemis degisiklikler kaydedilir.
        """
        app = get_app()
        doc = get_document(app, name)
        display_name = doc.DisplayName
        if save and doc.Dirty:
            doc.Save()
        doc.Close(True)
        return {"closed": display_name}

    @mcp.tool()
    @inventor_tool
    async def inventor_set_active_document(name: str) -> dict:
        """Belirtilen acik dokumani aktif (on plandaki) dokuman yapar."""
        app = get_app()
        doc = get_document(app, name)
        doc.Activate()
        return _doc_summary(doc)
