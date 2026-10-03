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
SERVER_INFO = {"name": "inventor", "version": "0.2.0"}

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


DOC = {"type": "string",
       "description": "Display name of an open document, as listed by `status` (e.g. \"Bracket.ipt\"). "
                      "Omit or leave empty to use the active document."}
BODY = {"type": "integer", "minimum": 1, "default": 1,
        "description": "1-based index of the solid/surface body in the part, as listed by `model_info`. Most parts have one body."}
READ_ONLY = {"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}

TOOLS = [
    {
        "name": "status",
        "title": "Inventor session status",
        "description": ("Checks that Autodesk Inventor is running and reachable over COM. Returns the Inventor version, "
                        "the active document, and every open document with its file path, type and unsaved-changes flag. "
                        "Call this first to get the document names the other tools accept."),
        "inputSchema": {"type": "object", "properties": {}},
        "annotations": READ_ONLY,
        "handler": t_status,
    },
    {
        "name": "model_info",
        "title": "Part bodies and features",
        "description": ("Summarizes a part document: for each body, whether it is solid, its face/edge/vertex counts and "
                        "its bounding box in mm; plus the feature tree (name, type, suppressed) and the number of Unwrap "
                        "features. Use it to orient yourself in a model before inspecting individual faces."),
        "inputSchema": {"type": "object", "properties": {"document": DOC}},
        "annotations": READ_ONLY,
        "handler": t_model_info,
    },
    {
        "name": "smallest_faces",
        "title": "Find the smallest faces",
        "description": ("Lists the N smallest faces of a body, sorted by area ascending. Each entry has the 1-based face "
                        "index, area in mm^2, edge count, surface type and center point in mm. Also returns the body's "
                        "total face count and total area. Use it to find import defects such as sliver or near-zero-area "
                        "faces from STEP/IGES files; the returned indices work with face_neighbors and unwrap_try."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document": DOC,
                "body": BODY,
                "count": {"type": "integer", "minimum": 1, "default": 15,
                          "description": "How many of the smallest faces to return."},
            },
        },
        "annotations": READ_ONLY,
        "handler": t_smallest_faces,
    },
    {
        "name": "face_neighbors",
        "title": "Faces adjacent to a face",
        "description": ("Returns the given face and every face that shares an edge with it, each with index, area in "
                        "mm^2, edge count, surface type and center point in mm; neighbors are sorted by area, largest "
                        "first. Use it to see what a suspicious small face sits between before excluding or fixing it."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document": DOC,
                "body": BODY,
                "index": {"type": "integer", "minimum": 1,
                          "description": "1-based index of the face to inspect, e.g. from smallest_faces."},
            },
            "required": ["index"],
        },
        "annotations": READ_ONLY,
        "handler": t_face_neighbors,
    },
    {
        "name": "unwrap_try",
        "title": "Dry-run an Unwrap",
        "description": ("Attempts Inventor's Unwrap feature on a set of faces and reports what happened: success or the "
                        "error and Inventor's last error message, time taken, and on success the result face count, "
                        "area in mm^2 and bounding box in mm. By default the created feature is deleted right away so "
                        "the document is left unchanged; pass keep=true to leave it in the model. Use it to find which "
                        "face set unwraps, which the Inventor UI does not tell you."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document": DOC,
                "body": BODY,
                "face_indices": {"type": "array", "items": {"type": "integer", "minimum": 1}, "minItems": 1,
                                 "description": "1-based indices of the faces to unwrap, as returned by smallest_faces or face_neighbors."},
                "auto_face_chain": {"type": "boolean", "default": False,
                                    "description": "Passed to Inventor's Unwrap definition as its auto face chain option."},
                "merge_result_body": {"type": "boolean", "default": False,
                                      "description": "Passed to Inventor's Unwrap definition as its merge result body option."},
                "alignment": {"type": "string", "enum": ["origin", "xy", "xz", "yz"], "default": "origin",
                              "description": "Passed to Inventor's Unwrap definition as its alignment option: origin, or the XY/XZ/YZ base plane."},
                "keep": {"type": "boolean", "default": False,
                         "description": "If true, the Unwrap feature stays in the document (modifies the model). "
                                        "If false (default), it is deleted after measuring."},
            },
            "required": ["face_indices"],
        },
        "annotations": {"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
        "handler": t_unwrap_try,
    },
    {
        "name": "script",
        "title": "Run PowerShell against Inventor",
        "description": ("Escape hatch for anything the other tools don't cover: runs arbitrary PowerShell with full "
                        "access to Inventor's COM API and the user's permissions. Predefined: $inv (Inventor.Application), "
                        "$MISSING (for optional COM arguments), Get-Doc <name> (document by display name, or active) and "
                        "FaceInfo <face> (area/center in mm). The script must write exactly one line of JSON, e.g. "
                        "`@{ n = $inv.Documents.Count } | ConvertTo-Json -Compress`. Prefer the dedicated tools when they "
                        "fit; this one can modify or close documents."),
        "inputSchema": {
            "type": "object",
            "properties": {
                "script": {"type": "string",
                           "description": "PowerShell code to run. Must output a single line of JSON."},
                "timeout": {"type": "integer", "minimum": 1, "default": 300,
                            "description": "Seconds to wait before the script is abandoned."},
            },
            "required": ["script"],
        },
        "annotations": {"readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False},
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
