#!/usr/bin/env python3
"""
Inventor MCP - a standalone MCP server with a live COM connection to Autodesk Inventor.

No external dependencies. Connects to Inventor through Windows PowerShell COM
(Marshal.GetActiveObject), so Inventor must already be running.

All lengths are returned in mm (Inventor's internal unit is cm; conversion happens
here). Areas are mm^2.
"""

import json
import subprocess
import sys
import textwrap

PROTOCOL_VERSION = "2024-11-05"
SERVER_INFO = {"name": "inventor", "version": "0.1.0"}

PS_EXE = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

# Prepended to every script: $inv (Application) and helpers are ready to use.
PREAMBLE = r"""
$ErrorActionPreference = "Stop"
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
  $inv = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Inventor.Application")
} catch {
  Write-Output (@{ error = "Inventor is not running or the COM connection failed: " + $_.Exception.Message } | ConvertTo-Json -Compress)
  exit 0
}
$MISSING = [System.Reflection.Missing]::Value

function Get-Doc($name) {
  if ([string]::IsNullOrEmpty($name)) { return $inv.ActiveDocument }
  foreach ($d in $inv.Documents) { if ($d.DisplayName -eq $name) { return $d } }
  throw "Document not found: $name"
}
function FaceInfo($f) {
  $r = $f.Evaluator.RangeBox
  return @{
    area_mm2 = [math]::Round($f.Evaluator.Area * 100, 4)
    edges    = $f.Edges.Count
    surface_type = [int]$f.SurfaceType
    cx_mm = [math]::Round(($r.MinPoint.X + $r.MaxPoint.X) * 5, 3)
    cy_mm = [math]::Round(($r.MinPoint.Y + $r.MaxPoint.Y) * 5, 3)
    cz_mm = [math]::Round(($r.MinPoint.Z + $r.MaxPoint.Z) * 5, 3)
  }
}
"""


def ps_str(value) -> str:
    """Quotes a value as a PowerShell single-quoted literal (no variable or subexpression expansion)."""
    return "'" + str(value).replace("'", "''") + "'"


def run_ps(body: str, timeout: int = 300) -> dict:
    """Runs a PowerShell script and parses its single-line JSON output."""
    script = PREAMBLE + "\n" + body
    try:
        proc = subprocess.run(
            [PS_EXE, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-Command", script],
            capture_output=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return {"error": f"PowerShell did not return within {timeout} s"}

    out = proc.stdout.decode("utf-8", errors="replace").strip()
    err = proc.stderr.decode("utf-8", errors="replace").strip()
    if not out:
        return {"error": err or f"no output (exit {proc.returncode})"}
    try:
        return json.loads(out)
    except json.JSONDecodeError:
        return {"raw": out, "stderr": err or None}


# --------------------------------------------------------------------------
# Tools
# --------------------------------------------------------------------------

def t_status(_args):
    return run_ps(r"""
$docs = @()
foreach ($d in $inv.Documents) {
  $docs += @{ name = $d.DisplayName; path = $d.FullFileName; type = [int]$d.DocumentType; dirty = [bool]$d.Dirty }
}
@{
  caption = $inv.Caption
  version = $inv.SoftwareVersion.DisplayVersion
  visible = [bool]$inv.Visible
  active  = if ($inv.ActiveDocument) { $inv.ActiveDocument.DisplayName } else { $null }
  documents = $docs
} | ConvertTo-Json -Depth 5 -Compress
""")


def t_model_info(args):
    doc = args.get("document", "")
    return run_ps(f"""
$doc = Get-Doc {ps_str(doc)}
$cd = $doc.ComponentDefinition
$bodies = @()
foreach ($b in $cd.SurfaceBodies) {{
  $bx = $b.RangeBox
  $bodies += @{{
    name = $b.Name; is_solid = [bool]$b.IsSolid
    faces = $b.Faces.Count; edges = $b.Edges.Count; vertices = $b.Vertices.Count
    bbox_mm = @{{
      xmin = [math]::Round($bx.MinPoint.X*10,3); xmax = [math]::Round($bx.MaxPoint.X*10,3)
      ymin = [math]::Round($bx.MinPoint.Y*10,3); ymax = [math]::Round($bx.MaxPoint.Y*10,3)
      zmin = [math]::Round($bx.MinPoint.Z*10,3); zmax = [math]::Round($bx.MaxPoint.Z*10,3)
    }}
  }}
}}
$feats = @()
foreach ($f in $cd.Features) {{ $feats += @{{ name = $f.Name; type = [int]$f.Type; suppressed = [bool]$f.Suppressed }} }}
@{{
  document = $doc.DisplayName; path = $doc.FullFileName; dirty = [bool]$doc.Dirty
  bodies = $bodies; features = $feats
  unwrap_features = $cd.Features.UnwrapFeatures.Count
}} | ConvertTo-Json -Depth 6 -Compress
""")


def t_smallest_faces(args):
    doc = args.get("document", "")
    n = int(args.get("count", 15))
    body = int(args.get("body", 1))
    return run_ps(f"""
$doc = Get-Doc {ps_str(doc)}
$b = $doc.ComponentDefinition.SurfaceBodies.Item({body})
$rows = @()
$i = 0
foreach ($f in $b.Faces) {{
  $i++
  $info = FaceInfo $f
  $info["index"] = $i
  $rows += $info
}}
$sorted = $rows | Sort-Object {{ $_.area_mm2 }} | Select-Object -First {n}
@{{
  total_faces = $rows.Count
  total_area_mm2 = [math]::Round((($rows | Measure-Object area_mm2 -Sum).Sum), 2)
  smallest = $sorted
}} | ConvertTo-Json -Depth 5 -Compress
""")


def t_face_neighbors(args):
    doc = args.get("document", "")
    body = int(args.get("body", 1))
    index = int(args["index"])
    return run_ps(f"""
$doc = Get-Doc {ps_str(doc)}
$b = $doc.ComponentDefinition.SurfaceBodies.Item({body})
$faces = @(); foreach ($f in $b.Faces) {{ $faces += $f }}
$target = $faces[{index} - 1]
$map = @{{}}
for ($i = 0; $i -lt $faces.Count; $i++) {{ $map[$faces[$i]] = $i + 1 }}
$nb = @{{}}
foreach ($e in $target.Edges) {{ foreach ($ff in $e.Faces) {{ if (-not $ff.Equals($target)) {{ $nb[$map[$ff]] = $ff }} }} }}
$rows = @()
foreach ($k in $nb.Keys) {{ $info = FaceInfo $nb[$k]; $info["index"] = $k; $rows += $info }}
$self = FaceInfo $target; $self["index"] = {index}
@{{ face = $self; neighbor_count = $rows.Count; neighbors = ($rows | Sort-Object {{ -$_.area_mm2 }}) }} | ConvertTo-Json -Depth 5 -Compress
""")


ALIGN = {"origin": 116993, "xy": 116994, "xz": 116995, "yz": 116996}


def t_unwrap_try(args):
    doc = args.get("document", "")
    body = int(args.get("body", 1))
    idx = args.get("face_indices") or []
    chain = bool(args.get("auto_face_chain", False))
    merge = bool(args.get("merge_result_body", False))
    keep = bool(args.get("keep", False))
    align = ALIGN.get(str(args.get("alignment", "origin")).lower(), 116993)
    idx_ps = ",".join(str(int(i)) for i in idx)
    return run_ps(f"""
$doc = Get-Doc {ps_str(doc)}
$cd = $doc.ComponentDefinition
$b = $cd.SurfaceBodies.Item({body})
$faces = @(); foreach ($f in $b.Faces) {{ $faces += $f }}
$sel = @({idx_ps})
$coll = $inv.TransientObjects.CreateFaceCollection()
foreach ($i in $sel) {{ $coll.Add($faces[$i - 1]) }}
$sw = [Diagnostics.Stopwatch]::StartNew()
try {{
  $def = $cd.Features.UnwrapFeatures.CreateDefinition($coll, $MISSING, {align}, $MISSING, $MISSING, ${'true' if chain else 'false'}, ${'true' if merge else 'false'})
  $uf = $cd.Features.UnwrapFeatures.Add($def)
  $sw.Stop()
  $a = 0; foreach ($ff in $uf.SurfaceBody.Faces) {{ $a += $ff.Evaluator.Area * 100 }}
  $rb = $uf.SurfaceBody.RangeBox
  $res = @{{
    ok = $true; feature = $uf.Name; seconds = [math]::Round($sw.Elapsed.TotalSeconds,2)
    input_faces = $sel.Count; result_faces = $uf.SurfaceBody.Faces.Count
    result_area_mm2 = [math]::Round($a,2)
    result_bbox_mm = @{{
      xmin=[math]::Round($rb.MinPoint.X*10,2); xmax=[math]::Round($rb.MaxPoint.X*10,2)
      ymin=[math]::Round($rb.MinPoint.Y*10,2); ymax=[math]::Round($rb.MaxPoint.Y*10,2)
      zmin=[math]::Round($rb.MinPoint.Z*10,2); zmax=[math]::Round($rb.MaxPoint.Z*10,2)
    }}
    kept = ${'true' if keep else 'false'}
  }}
  if (-not ${'true' if keep else 'false'}) {{ $uf.Delete() }}
  $res | ConvertTo-Json -Depth 5 -Compress
}} catch {{
  $sw.Stop()
  @{{ ok = $false; input_faces = $sel.Count; seconds = [math]::Round($sw.Elapsed.TotalSeconds,2)
     error = $_.Exception.Message
     inventor_message = $inv.ErrorManager.LastMessage }} | ConvertTo-Json -Depth 4 -Compress
}}
""", timeout=600)


def t_script(args):
    body = args["script"]
    return run_ps(body, timeout=int(args.get("timeout", 300)))


TOOLS = [
    {
        "name": "status",
        "description": "Connects to the running Inventor session; returns its version and open documents.",
        "inputSchema": {"type": "object", "properties": {}},
        "handler": t_status,
    },
    {
        "name": "model_info",
        "description": "Returns a part document's bodies (face/edge/vertex counts, bounding box in mm) and its feature list.",
        "inputSchema": {
            "type": "object",
            "properties": {"document": {"type": "string", "description": "Document display name; empty means the active document."}},
        },
        "handler": t_model_info,
    },
    {
        "name": "smallest_faces",
        "description": "Lists the smallest faces with area and center coordinates. Useful for finding import defects (slivers, tiny faces).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "document": {"type": "string"},
                "body": {"type": "integer", "default": 1},
                "count": {"type": "integer", "default": 15},
            },
        },
        "handler": t_smallest_faces,
    },
    {
        "name": "face_neighbors",
        "description": "Lists the faces that share an edge with the given face.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "document": {"type": "string"},
                "body": {"type": "integer", "default": 1},
                "index": {"type": "integer", "description": "1-based face index"},
            },
            "required": ["index"],
        },
        "handler": t_face_neighbors,
    },
    {
        "name": "unwrap_try",
        "description": ("Tries an Unwrap on the given faces. By default the feature is created and immediately deleted "
                        "(the document is not changed); keep=true leaves it in place. Returns the success/failure "
                        "details the GUI does not show."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document": {"type": "string"},
                "body": {"type": "integer", "default": 1},
                "face_indices": {"type": "array", "items": {"type": "integer"},
                                 "description": "1-based face indices"},
                "auto_face_chain": {"type": "boolean", "default": False},
                "merge_result_body": {"type": "boolean", "default": False},
                "alignment": {"type": "string", "enum": ["origin", "xy", "xz", "yz"], "default": "origin"},
                "keep": {"type": "boolean", "default": False,
                         "description": "if true, the created feature stays in the document"},
            },
            "required": ["face_indices"],
        },
        "handler": t_unwrap_try,
    },
    {
        "name": "script",
        "description": ("Runs arbitrary PowerShell against Inventor's COM API. $inv (Application), $MISSING and the "
                        "Get-Doc/FaceInfo helpers are predefined. The script must write single-line JSON "
                        "(ConvertTo-Json -Compress)."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "script": {"type": "string"},
                "timeout": {"type": "integer", "default": 300},
            },
            "required": ["script"],
        },
        "handler": t_script,
    },
]

HANDLERS = {t["name"]: t.pop("handler") for t in TOOLS}


# --------------------------------------------------------------------------
# JSON-RPC / MCP stdio loop
# --------------------------------------------------------------------------

def send(msg):
    sys.stdout.write(json.dumps(msg, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError:
            continue

        method = req.get("method")
        rid = req.get("id")

        if method == "initialize":
            asked = (req.get("params") or {}).get("protocolVersion") or PROTOCOL_VERSION
            send({"jsonrpc": "2.0", "id": rid, "result": {
                "protocolVersion": asked,
                "capabilities": {"tools": {}},
                "serverInfo": SERVER_INFO,
            }})
        elif method == "notifications/initialized":
            pass
        elif method == "ping":
            send({"jsonrpc": "2.0", "id": rid, "result": {}})
        elif method == "tools/list":
            send({"jsonrpc": "2.0", "id": rid, "result": {"tools": TOOLS}})
        elif method == "tools/call":
            params = req.get("params") or {}
            name = params.get("name")
            args = params.get("arguments") or {}
            fn = HANDLERS.get(name)
            if fn is None:
                send({"jsonrpc": "2.0", "id": rid,
                      "error": {"code": -32601, "message": f"unknown tool: {name}"}})
                continue
            try:
                result = fn(args)
                is_err = isinstance(result, dict) and "error" in result
                send({"jsonrpc": "2.0", "id": rid, "result": {
                    "content": [{"type": "text",
                                 "text": json.dumps(result, ensure_ascii=False, indent=2)}],
                    "isError": bool(is_err),
                }})
            except Exception as exc:  # a failing tool must not take the server down
                send({"jsonrpc": "2.0", "id": rid, "result": {
                    "content": [{"type": "text", "text": f"{type(exc).__name__}: {exc}"}],
                    "isError": True,
                }})
        elif rid is not None:
            send({"jsonrpc": "2.0", "id": rid,
                  "error": {"code": -32601, "message": f"unsupported method: {method}"}})


if __name__ == "__main__":
    main()
