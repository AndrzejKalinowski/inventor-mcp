# Autodesk Inventor MCP

An [MCP](https://modelcontextprotocol.io) server that bridges Claude (or any MCP-compatible client) to **Autodesk Inventor** via COM automation. It exposes Inventor's part modeling, assembly, drawing, and document tools as callable MCP tools, so an AI assistant can sketch, extrude, fillet, place assembly components, generate drawing views, and more — directly inside a running Inventor session.

> **Platform note:** Autodesk Inventor's COM API is Windows-only. This server must run on Windows with Inventor installed, using a Python environment with `pywin32`.

## Features

- **Sketching** — lines, rectangles, circles, arcs, polygons, polylines, slots, ellipses, splines, fillets, construction geometry, mirroring, constraints, and dimensions.
- **3D part features** — extrude, revolve, sweep, loft, coil (helical threads/springs), hole, fillet, chamfer, shell, rectangular/circular pattern, mirror, combine (boolean union/subtract/intersect).
- **Work geometry** — work planes, axes, and points.
- **Parameters** — list, read, set, and create model parameters.
- **Properties** — iProperties and mass properties (volume, mass, center of gravity, etc.).
- **Geometry inspection** — list bodies, faces, edges, and bounding boxes.
- **Assemblies** — place/ground/position components, add and inspect mate constraints, rectangular/circular component patterns, bill of materials.
- **Drawings** — sheets, base/projected views, dimensions, parts lists, title block fields.
- **Appearances** — list and apply visual styles/materials.
- **Document & app control** — new/open/save/close documents, active document switching, status checks.
- **Visualization** — zoom-to-fit and screenshot capture for visual verification.

Tool implementations live under `inventor/tools/`, organized by domain (`sketch_tools.py`, `feature_tools.py`, `assembly_tools.py`, `drawing_tools.py`, `geometry_tools.py`, `parameter_tools.py`, `property_tools.py`, `appearance_tools.py`, `export_tools.py`, `view_tools.py`, `app_tools.py`).

## Requirements

- Windows with **Autodesk Inventor** installed and licensed.
- Python 3.10+ (a virtual environment is recommended).
- Dependencies listed in `requirements.txt`, notably:
  - [`mcp`](https://pypi.org/project/mcp/) (FastMCP server framework)
  - [`pywin32`](https://pypi.org/project/pywin32/) (COM automation)

## Installation

```powershell
# from the repository root
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Running the server

The server communicates over stdio, as expected by MCP clients such as Claude Desktop:

```powershell
.venv\Scripts\python.exe inventor_mcp.py
```

It will connect to a running Inventor instance, or launch one if none is open.

### Using with Claude Desktop / Cowork

Add an entry to your MCP client's configuration pointing at the venv interpreter and `inventor_mcp.py`, for example:

```json
{
  "mcpServers": {
    "inventor": {
      "command": "C:\\path\\to\\repo\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\repo\\inventor_mcp.py"]
    }
  }
}
```

A ready-to-customize Claude Desktop Extension (`.mcpb`) manifest template is provided at `inventor-mcp-ext/manifest.example.json` — copy it to `inventor-mcp-ext/manifest.json`, replace the `<ABSOLUTE_PATH_TO_REPO>` placeholders with your local checkout path, then package or load it.

## Architecture notes

- `inventor/connection.py` manages the connection to `Inventor.Application`, re-initializing COM per call to stay safe across the multi-threaded calls an MCP server can receive.
- `inventor/registry.py` tracks sketch entities and other COM object references by an opaque `entity_id` string, so tool calls can refer back to objects created in earlier calls without re-querying geometry.
- `inventor/helpers.py` and `inventor/units.py` provide shared geometry helpers and millimeter/centimeter unit conversions (Inventor's API is internally in centimeters).
- Several feature/sketch collections in Inventor's COM type library are wrapped with `win32com.client.dynamic.Dispatch` rather than the early-bound gencache proxy, working around a `gen_py` cache staleness issue observed with certain collections (e.g. `CoilFeatures`, `SweepFeatures`, `LoftFeatures`, `CombineFeatures`).
- Closed sketch profiles are built by chaining each new line from the previous line's own endpoint object, rather than via `AddCoincident` on independent points — this avoids a COM error Inventor raises when merging two line-endpoint `SketchPoint`s directly.

## License

Free for non-commercial use. Commercial use requires prior written permission from the author — see [LICENSE](LICENSE).
