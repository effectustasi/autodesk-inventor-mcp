# inventor-mcp

[![effectustasi/autodesk-inventor-mcp MCP server](https://glama.ai/mcp/servers/effectustasi/autodesk-inventor-mcp/badges/score.svg)](https://glama.ai/mcp/servers/effectustasi/autodesk-inventor-mcp)

An [MCP](https://modelcontextprotocol.io) server that gives AI agents (Claude Code, Claude Desktop, Cursor, Codex…) a **live connection to Autodesk Inventor**.

Ask your agent things like *"which faces in this imported STEP are tiny slivers?"* or *"will Unwrap work on these 12 faces?"* and it reads the answer straight from the open model.

- **Zero dependencies:** one Python file, standard library only.
- **Live:** talks to the running Inventor session over COM (via Windows PowerShell). No export, no file round-trips.
- **Units you think in:** lengths in mm, areas in mm² (Inventor's internal unit is cm; the server converts).
- **Safe by default:** `unwrap_try` creates the feature, measures the result and deletes it again unless you pass `keep: true`.

## Requirements

- Windows with Autodesk Inventor installed and **running** (developed against Inventor 2027)
- Python 3.9+

## Install

Clone or download `server.py`, then register it with your agent.

**Claude Code**

```bash
claude mcp add inventor -- python C:\path\to\inventor-mcp\server.py
```

**Claude Desktop / Cursor / any MCP client** (`mcpServers` config):

```json
{
  "mcpServers": {
    "inventor": {
      "command": "python",
      "args": ["C:\\path\\to\\inventor-mcp\\server.py"]
    }
  }
}
```

## Tools

| Tool | What it does |
|---|---|
| `status` | Inventor version and open documents |
| `model_info` | Bodies (face/edge/vertex counts, bounding box in mm) and the feature tree of a part |
| `smallest_faces` | The N smallest faces with area and center. Finds import defects fast |
| `face_neighbors` | Faces that share an edge with a given face |
| `unwrap_try` | Tries an Unwrap on a face set and reports success, timing, result area and bounding box, the details the GUI doesn't show. Rolls back by default |
| `script` | Runs arbitrary PowerShell against Inventor's COM API, with `$inv` and helpers predefined |

Face indices are 1-based and stable for a given body as long as the model doesn't change.

## ⚠️ About the `script` tool

`script` executes **any PowerShell** your agent writes, with your user's permissions. That's what makes it powerful (anything in the Inventor API is reachable), and it's also why you should keep your agent's tool-approval prompts on for it. If you don't want it, delete its entry from `TOOLS` in `server.py`.

All other tools pass arguments as data, never as code.

## Example (illustrative)

> **You:** The STEP I imported won't unwrap. What's wrong with it?
>
> **Agent:** *(calls `smallest_faces`)* Body 1 has 412 faces; the 6 smallest are under 0.01 mm², clustered around (12.4, -3.1, 40.0). *(calls `face_neighbors` on face 287)* They sit between two large B-spline faces, which is the classic sliver from a bad tessellation export. *(calls `unwrap_try` without them)* Unwrap succeeds in 1.8 s when those 6 faces are excluded.

## Contributing

Issues and PRs are welcome, especially new read-only tools (assemblies, sketches, iProperties, sheet metal), Inventor-version compatibility reports, and non-English Windows fixes. See the issues labeled `good first issue`.

## License

MIT. Not affiliated with or endorsed by Autodesk. Autodesk and Inventor are trademarks of Autodesk, Inc.

## More tools by effectustasi

- [agent-receipts](https://github.com/effectustasi/agent-receipts): Skills that make AI coding agents prove "done" with real test output
- [blender-dlss5-neural-rendering](https://github.com/effectustasi/blender-dlss5-neural-rendering): Blender viewport and renders through DLSS 5 neural rendering
- [metahuman-face-capture](https://github.com/effectustasi/metahuman-face-capture): MetaHuman face capture from a webcam in Blender
- [unreal-groom-alembic-exporter](https://github.com/effectustasi/unreal-groom-alembic-exporter): Export UE Groom assets (MetaHuman hair) to Alembic
