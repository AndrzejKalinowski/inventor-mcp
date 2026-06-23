"""Parametre araclari: model/kullanici parametrelerini listeleme, guncelleme, olusturma."""

import pywintypes
from win32com.client import constants

from inventor.connection import get_app
from inventor.helpers import get_document, get_part_component_def, inventor_tool

PARAMETER_TYPE_NAMES = {
    constants.kModelParameter: "model",
    constants.kReferenceParameter: "reference",
    constants.kTableParameter: "table",
    constants.kUserParameter: "user",
    constants.kDerivedParameter: "derived",
    constants.kFinishParameter: "finish",
}

# AddByExpression icin denenecek UnitsSpecifier sirasi: uzunluk -> aci -> birimsiz.
_UNITS_SPECIFIER_CANDIDATES = [
    constants.kDatabaseLengthUnits,
    constants.kDatabaseAngleUnits,
    constants.kUnitlessUnits,
]


def _parameter_summary(p) -> dict:
    return {
        "name": p.Name,
        "value": p.Value,
        "expression": p.Expression,
        "units": p.Units,
        "comment": p.Comment,
        "kind": PARAMETER_TYPE_NAMES.get(p.ParameterType, str(p.ParameterType)),
    }


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_list_parameters(doc: str | None = None) -> list[dict]:
        """Part dokumanindaki tum parametreleri (model + kullanici + referans) listeler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        return [_parameter_summary(p) for p in comp_def.Parameters]

    @mcp.tool()
    @inventor_tool
    async def inventor_set_parameter(name: str, expression: str, doc: str | None = None) -> dict:
        """Var olan bir parametrenin ifadesini (expression) degistirir.

        expression: '12 mm', '45 deg', '2 * d0' gibi birim/ifade iceren bir string.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        try:
            param = comp_def.Parameters.Item(name)
        except pywintypes.com_error:
            raise LookupError(f"'{name}' adli parametre bulunamadi.")
        param.Expression = expression
        return _parameter_summary(param)

    @mcp.tool()
    @inventor_tool
    async def inventor_create_parameter(
        name: str, expression: str, comment: str | None = None, doc: str | None = None
    ) -> dict:
        """Yeni bir kullanici parametresi (UserParameter) olusturur.

        expression: '12 mm', '45 deg', '2' gibi birim/ifade iceren bir string.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        user_params = comp_def.Parameters.UserParameters

        param = None
        last_error = None
        for units_specifier in _UNITS_SPECIFIER_CANDIDATES:
            try:
                param = user_params.AddByExpression(name, expression, units_specifier)
                break
            except pywintypes.com_error as e:
                last_error = e
        if param is None:
            raise last_error

        if comment:
            param.Comment = comment
        return _parameter_summary(param)
