"""Goruntu ve ekran goruntusu araclari."""

from inventor.connection import get_app
from inventor.helpers import get_document, inventor_tool


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_zoom_all(doc: str | None = None) -> dict:
        """Aktif goruntude tum modeli ekrana sigdirir (zoom to fit)."""
        app = get_app()
        document = get_document(app, doc)
        document.Activate()
        view = app.ActiveView
        view.Fit(True)
        view.Update()
        return {"doc": document.DisplayName, "zoomed": True}

    @mcp.tool()
    @inventor_tool
    async def inventor_screenshot(path: str, width: int = 800, height: int = 600, doc: str | None = None) -> dict:
        """Aktif goruntunun ekran goruntusunu PNG olarak kaydeder.

        Onceden tum modeli ekrana sigdirir (zoom to fit).
        """
        app = get_app()
        document = get_document(app, doc)
        document.Activate()
        view = app.ActiveView
        view.Fit(True)
        view.Update()
        view.SaveAsBitmap(path, width, height)
        return {"path": path, "width": width, "height": height}
