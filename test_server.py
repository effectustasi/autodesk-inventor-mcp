"""Smoke test: the server starts, speaks MCP over stdio, and lists well-formed tools.

Runs without Inventor (and without Windows): tool calls are expected to fail cleanly.
Usage: python test_server.py
"""

import json
import subprocess
import sys

EXPECTED = {"status", "model_info", "smallest_faces", "face_neighbors", "unwrap_try", "script"}

requests = [
    {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-06-18"}},
    {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
    {"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "status", "arguments": {}}},
    {"jsonrpc": "2.0", "id": 4, "method": "tools/call", "params": {"name": "no_such_tool", "arguments": {}}},
]
proc = subprocess.run(
    [sys.executable, "server.py"],
    input="\n".join(json.dumps(r) for r in requests) + "\n",
    capture_output=True, text=True, encoding="utf-8", timeout=120,
)
replies = {m["id"]: m for m in map(json.loads, proc.stdout.splitlines())}

assert replies[1]["result"]["protocolVersion"] == "2025-06-18", replies[1]

tools = replies[2]["result"]["tools"]
assert {t["name"] for t in tools} == EXPECTED, [t["name"] for t in tools]
for t in tools:
    assert "handler" not in t, t["name"]
    assert len(t["description"]) > 40, t["name"]
    for name, prop in t["inputSchema"].get("properties", {}).items():
        assert prop.get("description"), f"{t['name']}.{name} has no description"

# Inventor is not running in CI: the call must come back as a tool error, not crash the server.
assert replies[3]["result"]["isError"] is True, replies[3]
assert replies[4]["error"]["code"] == -32601, replies[4]

print(f"OK: {len(tools)} tools, all parameters documented, errors handled")
