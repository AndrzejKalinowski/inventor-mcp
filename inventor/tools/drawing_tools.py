"""Cizim (drawing) araclari: sayfalar, gorunumler, olculendirme, parca listesi, antet."""

import pywintypes
from win32com.client import constants

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import get_document, get_drawing_document, get_sheet, inventor_tool, point2d
from inventor.units import cm_to_mm

VIEW_ORIENTATIONS = {
    "front": constants.kFrontViewOrientation,
    "back": constants.kBackViewOrientation,
    "top": constants.kTopViewOrientation,
    "bottom": constants.kBottomViewOrientation,
    "left": constants.kLeftViewOrientation,
    "right": constants.kRightViewOrientation,
    "iso_top_right": constants.kIsoTopRightViewOrientation,
    "iso_top_left": constants.kIsoTopLeftViewOrientation,
    "iso_bottom_right": constants.kIsoBottomRightViewOrientation,
    "iso_bottom_left": constants.kIsoBottomLeftViewOrientation,
}

VIEW_STYLES = {
    "hidden_line": constants.kHiddenLineDrawingViewStyle,
    "hidden_line_removed": constants.kHiddenLineRemovedDrawingViewStyle,
    "shaded": constants.kShadedDrawingViewStyle,
}

CURVE_TYPE_NAMES = {
    constants.kLineSegmentCurve: "line",
    constants.kCircleCurve: "circle",
    constants.kCircularArcCurve: "arc",
    constants.kEllipseFullCurve: "ellipse",
    constants.kEllipticalArcCurve: "elliptical_arc",
}

INTENT_POINTS = {
    "start": constants.kStartPointIntent,
    "end": constants.kEndPointIntent,
    "center": constants.kCenterPointIntent,
    "mid": constants.kMidPointIntent,
}

DIM_ORIENTATIONS = {
    "aligned": constants.kAlignedDimensionType,
    "horizontal": constants.kHorizontalDimensionType,
    "vertical": constants.kVerticalDimensionType,
}

DIM_TYPE_NAMES = {
    constants.kLinearGeneralDimensionObject: "linear",
    constants.kRadiusGeneralDimensionObject: "radius",
    constants.kDiameterGeneralDimensionObject: "diameter",
    constants.kAngularGeneralDimensionObject: "angular",
}


def _sheet_summary(sheet) -> dict:
    return {
        "name": sheet.Name,
        "width_mm": round(cm_to_mm(sheet.Width), 2),
        "height_mm": round(cm_to_mm(sheet.Height), 2),
        "view_count": sheet.DrawingViews.Count,
    }


def _view_summary(view) -> dict:
    entry = {
        "name": view.Name,
        "left_mm": round(cm_to_mm(view.Left), 4),
        "top_mm": round(cm_to_mm(view.Top), 4),
        "width_mm": round(cm_to_mm(view.Width), 4),
        "height_mm": round(cm_to_mm(view.Height), 4),
        "scale": view.Scale,
    }
    try:
        entry["referenced_document"] = view.ReferencedDocumentDescriptor.FullDocumentName
    except pywintypes.com_error:
        entry["referenced_document"] = None
    return entry


def _curve_summary(curve) -> dict:
    kind = CURVE_TYPE_NAMES.get(curve.CurveType, str(curve.CurveType))
    entry = {"kind": kind}
    sp, ep = curve.StartPoint, curve.EndPoint
    if sp is not None and ep is not None:
        entry["start_mm"] = [round(cm_to_mm(sp.X), 4), round(cm_to_mm(sp.Y), 4)]
        entry["end_mm"] = [round(cm_to_mm(ep.X), 4), round(cm_to_mm(ep.Y), 4)]
    if kind in ("circle", "arc", "ellipse", "elliptical_arc"):
        cp = curve.CenterPoint
        entry["center_mm"] = [round(cm_to_mm(cp.X), 4), round(cm_to_mm(cp.Y), 4)]
    return entry


def _dimension_summary(dim) -> dict:
    return {"value_mm": round(cm_to_mm(dim.ModelValue), 4), "text": dim.Text.Text}


def _resolve_view(sheet, name: str):
    views = sheet.DrawingViews
    for i in range(1, views.Count + 1):
        v = views.Item(i)
        if v.Name == name:
            return v
    raise LookupError(f"'{name}' adli gorunum bulunamadi.")


def _resolve_intent(point: str) -> int:
    intent = INTENT_POINTS.get(point.lower())
    if intent is None:
        raise LookupError(f"Gecersiz nokta turu: '{point}'. Gecerli degerler: {', '.join(INTENT_POINTS)}.")
    return intent


def _parts_list_summary(pl) -> dict:
    columns = [pl.PartsListColumns.Item(i).Title for i in range(1, pl.PartsListColumns.Count + 1)]
    rows = []
    for i in range(1, pl.PartsListRows.Count + 1):
        row = pl.PartsListRows.Item(i)
        rows.append({col: row.Item(col).Value for col in columns})
    return {"columns": columns, "rows": rows}


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_list_sheets(doc: str | None = None) -> list[dict]:
        """Drawing dokumanindaki tum sayfalari (sheet) listeler."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sheets = dwg_doc.Sheets
        return [_sheet_summary(sheets.Item(i)) for i in range(1, sheets.Count + 1)]

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_add_sheet(name: str | None = None, doc: str | None = None) -> dict:
        """Drawing dokumanina yeni bir sayfa (sheet) ekler."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        new_sheet = dwg_doc.Sheets.Add(SheetName=name) if name else dwg_doc.Sheets.Add()
        return _sheet_summary(new_sheet)

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_delete_sheet(name: str, doc: str | None = None) -> dict:
        """Adi verilen sayfayi (sheet) drawing dokumanindan siler."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, name)
        sht.Delete()
        return {"deleted": name}

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_add_base_view(
        model_doc: str,
        x_mm: float,
        y_mm: float,
        scale: float = 1.0,
        orientation: str = "front",
        style: str = "hidden_line",
        sheet: str | None = None,
        doc: str | None = None,
    ) -> dict:
        """Acik bir parca/montaj dokumanindan sayfaya bir temel (base) gorunum ekler.

        model_doc: gorunumu eklenecek acik dokumanin adi; bu dokuman diske
        kaydedilmis olmalidir (FullFileName bos olmamali).
        orientation: 'front', 'back', 'top', 'bottom', 'left', 'right',
        'iso_top_right', 'iso_top_left', 'iso_bottom_right', 'iso_bottom_left'.
        style: 'hidden_line', 'hidden_line_removed', 'shaded'.
        x_mm/y_mm: sayfa uzerindeki konum (sol-alt kosegen mm).

        Not: kaydedilmemis bir drawing dokumaninda ilk base view eklendiginde
        Inventor dokumanin adini (DisplayName) referans alinan modelin adina
        gore degistirebilir. Donen sonuctaki 'doc' alani, sonraki cagrilarda
        kullanilmasi gereken guncel dokuman adidir.
        """
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)

        model = get_document(app, model_doc)
        if not model.FullFileName:
            raise LookupError(
                f"'{model.DisplayName}' diske kaydedilmemis. Once inventor_save_document ile kaydedin."
            )

        orient = VIEW_ORIENTATIONS.get(orientation.lower())
        if orient is None:
            raise LookupError(
                f"Gecersiz orientation: '{orientation}'. Gecerli degerler: {', '.join(VIEW_ORIENTATIONS)}."
            )
        view_style = VIEW_STYLES.get(style.lower())
        if view_style is None:
            raise LookupError(f"Gecersiz style: '{style}'. Gecerli degerler: {', '.join(VIEW_STYLES)}.")

        pos = point2d(app, x_mm, y_mm)
        view = sht.DrawingViews.AddBaseView(model, pos, scale, orient, view_style)
        return {"doc": document.DisplayName, **_view_summary(view)}

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_add_projected_view(
        base_view: str,
        x_mm: float,
        y_mm: float,
        style: str | None = None,
        sheet: str | None = None,
        doc: str | None = None,
    ) -> dict:
        """Var olan bir gorunumden projeksiyon (projected view) olusturur.

        style verilmezse kaynak gorunumun stili kullanilir.
        """
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        parent = _resolve_view(sht, base_view)

        if style:
            view_style = VIEW_STYLES.get(style.lower())
            if view_style is None:
                raise LookupError(f"Gecersiz style: '{style}'. Gecerli degerler: {', '.join(VIEW_STYLES)}.")
        else:
            view_style = parent.ViewStyle

        pos = point2d(app, x_mm, y_mm)
        view = sht.DrawingViews.AddProjectedView(parent, pos, view_style)
        return _view_summary(view)

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_list_views(sheet: str | None = None, doc: str | None = None) -> list[dict]:
        """Sayfadaki tum gorunumleri listeler."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        views = sht.DrawingViews
        return [_view_summary(views.Item(i)) for i in range(1, views.Count + 1)]

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_set_view_position(
        view: str, x_mm: float, y_mm: float, sheet: str | None = None, doc: str | None = None
    ) -> dict:
        """Bir gorunumun sayfa uzerindeki konumunu degistirir (mm)."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        v = _resolve_view(sht, view)
        v.Position = point2d(app, x_mm, y_mm)
        return _view_summary(v)

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_delete_view(view: str, sheet: str | None = None, doc: str | None = None) -> dict:
        """Bir gorunumu sayfadan siler."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        v = _resolve_view(sht, view)
        name = v.Name
        v.Delete()
        return {"deleted": name}

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_list_view_curves(view: str, sheet: str | None = None, doc: str | None = None) -> dict:
        """Bir gorunumdeki tum kenar izdusumlerini (DrawingCurve) listeler.

        Donen entity_id'ler `inventor_drawing_add_dimension` icin kullanilabilir.
        kind: 'line', 'circle', 'arc'. start_mm/end_mm/center_mm sayfa uzerindeki
        konumlardir (mm), olculendirme degeri modelden okunur.
        """
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        v = _resolve_view(sht, view)

        group = f"view:{sht.Name}:{v.Name}"
        registry.clear_sketch(document.DisplayName, group)

        curves = v.GetDrawingCurves()
        result = []
        for i in range(1, curves.Count + 1):
            c = curves.Item(i)
            entity_id = registry.register(document.DisplayName, group, "curve", c)
            entry = {"entity_id": entity_id, "index": i}
            entry.update(_curve_summary(c))
            result.append(entry)
        return {"view": v.Name, "curves": result}

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_add_dimension(
        type: str,
        entity_ids: list[str],
        points: list[str] | None = None,
        x_mm: float = 0.0,
        y_mm: float = 0.0,
        orientation: str = "aligned",
        sheet: str | None = None,
        doc: str | None = None,
    ) -> dict:
        """Sayfaya bir olculendirme (DrawingDimension) ekler ve modelden okunan degerini dondurur.

        type: 'linear' (1 veya 2 entity_id), 'radius' (1 entity_id, daire/yay),
        'diameter' (1 entity_id, daire).
        entity_ids/points: `inventor_drawing_list_view_curves`'dan alinan curve
        entity_id'leri ve her biri icin nokta turu ('start', 'end', 'center', 'mid').
        - linear: 1 entity_id verilirse points varsayilan ['start','end'] (o
          kenarin uzunlugu); 2 entity_id verilirse her biri icin bir nokta gerekir.
        - radius/diameter: points varsayilan ['center'].
        orientation: linear icin 'aligned', 'horizontal', 'vertical'.
        x_mm/y_mm: olculendirme yazisinin sayfa uzerindeki konumu (mm).
        """
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        text_pos = point2d(app, x_mm, y_mm)

        kind = type.lower()
        curves = [registry.resolve(eid) for eid in entity_ids]

        if kind == "linear":
            if len(curves) == 1:
                curves = [curves[0], curves[0]]
            if len(curves) != 2:
                raise LookupError("'linear' boyutu 1 veya 2 entity_id gerektirir.")
            pts = points or ["start", "end"]
            if len(pts) != 2:
                raise LookupError("'linear' boyutu icin 'points' tam olarak 2 eleman icermeli.")
            intent_one = sht.CreateGeometryIntent(curves[0], _resolve_intent(pts[0]))
            intent_two = sht.CreateGeometryIntent(curves[1], _resolve_intent(pts[1]))
            dim_orient = DIM_ORIENTATIONS.get(orientation.lower())
            if dim_orient is None:
                raise LookupError(
                    f"Gecersiz orientation: '{orientation}'. Gecerli degerler: {', '.join(DIM_ORIENTATIONS)}."
                )
            dim = sht.DrawingDimensions.GeneralDimensions.AddLinear(text_pos, intent_one, intent_two, dim_orient)
        elif kind in ("radius", "diameter"):
            if len(curves) != 1:
                raise LookupError(f"'{kind}' boyutu 1 entity_id gerektirir.")
            pts = points or ["center"]
            intent = sht.CreateGeometryIntent(curves[0], _resolve_intent(pts[0]))
            if kind == "diameter":
                dim = sht.DrawingDimensions.GeneralDimensions.AddDiameter(text_pos, intent)
            else:
                dim = sht.DrawingDimensions.GeneralDimensions.AddRadius(text_pos, intent)
        else:
            raise LookupError(f"Gecersiz type: '{type}'. Gecerli turler: linear, radius, diameter.")

        return {"type": kind, **_dimension_summary(dim)}

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_list_dimensions(sheet: str | None = None, doc: str | None = None) -> list[dict]:
        """Sayfadaki tum olculendirmeleri (deger ve metin ile) listeler."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        dims = sht.DrawingDimensions
        result = []
        for i in range(1, dims.Count + 1):
            dim = dims.Item(i)
            entry = {"index": i, "type": DIM_TYPE_NAMES.get(dim.Type, str(dim.Type))}
            try:
                entry.update(_dimension_summary(dim))
            except pywintypes.com_error:
                pass
            result.append(entry)
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_add_parts_list(
        view: str, x_mm: float, y_mm: float, sheet: str | None = None, doc: str | None = None
    ) -> dict:
        """Bir montaj gorunumune dayali parca listesi (parts list) ekler.

        x_mm/y_mm: parca listesinin sol-ust kosesinin sayfa uzerindeki konumu (mm).
        """
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        v = _resolve_view(sht, view)
        pos = point2d(app, x_mm, y_mm)
        pl = sht.PartsLists.Add(v, pos)
        return _parts_list_summary(pl)

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_get_parts_list(sheet: str | None = None, doc: str | None = None) -> dict:
        """Sayfadaki mevcut parca listesinin sutunlarini ve satirlarini dondurur."""
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        pls = sht.PartsLists
        if pls.Count == 0:
            raise LookupError("Bu sayfada bir parca listesi yok.")
        return _parts_list_summary(pls.Item(1))

    @mcp.tool()
    @inventor_tool
    async def inventor_drawing_list_title_block_fields(sheet: str | None = None, doc: str | None = None) -> list[dict]:
        """Sayfanin antet (title block) alanlarini ve guncel degerlerini listeler.

        'definition' alaninda bir <Property PropertySet='...' Property='...'>
        veya <DerivedProperty> referansi varsa, bu alanin degeri ilgili iProperty'den
        gelir; degerini degistirmek icin `inventor_set_iproperty(propset, name, value)`
        kullanin (PropertySet/Property adlari 'definition' icinden okunabilir).
        Statik etiketler (orn. 'TITLE', 'SCALE') ve <StyleOverride> ile sarili duz
        metinler sabittir, degistirilemez.
        """
        app = get_app()
        document = get_document(app, doc)
        dwg_doc = get_drawing_document(document)
        sht = get_sheet(dwg_doc, sheet)
        tb = sht.TitleBlock
        texts = tb.Definition.Sketch.TextBoxes

        result = []
        for i in range(1, texts.Count + 1):
            t = texts.Item(i)
            result.append({"index": i, "definition": t.FormattedText, "value": tb.GetResultText(t)})
        return result
