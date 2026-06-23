"""Autodesk Inventor MCP koprusu - giris noktasi.

Claude Desktop bu dosyayi (venv python'i ile) stdio MCP server olarak calistirir.
"""

import pythoncom

pythoncom.CoInitialize()

from mcp.server.fastmcp import FastMCP

from inventor.tools import (
    app_tools,
    appearance_tools,
    assembly_tools,
    drawing_tools,
    export_tools,
    feature_tools,
    geometry_tools,
    parameter_tools,
    property_tools,
    sketch_tools,
    view_tools,
)

mcp = FastMCP("inventor")

app_tools.register(mcp)
sketch_tools.register(mcp)
feature_tools.register(mcp)
parameter_tools.register(mcp)
property_tools.register(mcp)
geometry_tools.register(mcp)
assembly_tools.register(mcp)
drawing_tools.register(mcp)
appearance_tools.register(mcp)
export_tools.register(mcp)
view_tools.register(mcp)


if __name__ == "__main__":
    mcp.run()
