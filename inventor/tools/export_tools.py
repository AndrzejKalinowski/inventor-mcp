"""STEP/IGES/STL/SAT/DWG/PDF disa aktarim araclari."""

import os

from win32com.client import CastTo, constants

from inventor.connection import get_app
from inventor.helpers import get_document, inventor_tool

TRANSLATOR_IDS = {
    "step": "{90AF7F40-0C01-11D5-8E83-0010B541CD80}",
    "iges": "{90AF7F44-0C01-11D5-8E83-0010B541CD80}",
    "stl": "{533E9A98-FC3B-11D4-8E7E-0010B541CD80}",
    "sat": "{89162634-02B6-11D5-8E80-0010B541CD80}",
    "dwg": "{C24E3AC2-122E-11D5-8E91-0010B541CD80}",
    "pdf": "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}",
}


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_export(format: str, path: str, doc: str | None = None) -> dict:
        """Dokumani belirtilen formatta disa aktarir (Save Copy As).

        format: 'step', 'iges', 'stl', 'sat', 'dwg', 'pdf'.
        path: cikti dosyasinin tam yolu.
        """
        app = get_app()
        document = get_document(app, doc)

        clsid = TRANSLATOR_IDS.get(format.lower())
        if clsid is None:
            raise LookupError(f"Gecersiz format: '{format}'. Gecerli degerler: {', '.join(TRANSLATOR_IDS)}.")

        addin = app.ApplicationAddIns.ItemById(clsid)
        translator = CastTo(addin, "TranslatorAddIn")

        context = app.TransientObjects.CreateTranslationContext()
        context.Type = constants.kFileBrowseIOMechanism
        options = app.TransientObjects.CreateNameValueMap()
        data_medium = app.TransientObjects.CreateDataMedium()
        data_medium.FileName = path

        translator.HasSaveCopyAsOptions(document, context, options)
        translator.SaveCopyAs(document, context, options, data_medium)

        return {"format": format.lower(), "path": path, "exists": os.path.exists(path)}
