#!/usr/bin/env python3
"""Helper to call Zotero MCP server via Streamable HTTP."""
import json, sys, urllib.request

try:
    import config
    URL = config.get_mcp_url()
except Exception:
    URL = "http://127.0.0.1:23120/mcp"
_session_id = None

def _post(payload, extra_headers=None):
    headers = {"Content-Type": "application/json", "Accept": "application/json, text/event-stream"}
    if _session_id:
        headers["Mcp-Session-Id"] = _session_id
    if extra_headers:
        headers.update(extra_headers)
    data = json.dumps(payload).encode()
    req = urllib.request.Request(URL, data=data, headers=headers, method="POST")
    try:
        resp = urllib.request.urlopen(req, timeout=60)
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        return e.code, dict(e.headers), body
    return resp.status, dict(resp.headers), resp.read().decode()

def _parse(resp_body):
    """Parse SSE or plain JSON response, return the result object."""
    body = resp_body.strip()
    if body.startswith("event:") or body.startswith("data:"):
        # SSE: find the data: line with the JSON-RPC response
        for line in body.splitlines():
            if line.startswith("data:"):
                payload = line[5:].strip()
                if payload:
                    try:
                        obj = json.loads(payload)
                        if "result" in obj:
                            return obj["result"]
                        if "error" in obj:
                            return {"error": obj["error"]}
                    except json.JSONDecodeError:
                        pass
        return {"raw_sse": body}
    else:
        try:
            obj = json.loads(body)
            if "result" in obj:
                return obj["result"]
            if "error" in obj:
                return {"error": obj["error"]}
            return obj
        except json.JSONDecodeError:
            return {"raw": body}

def initialize():
    global _session_id
    payload = {"jsonrpc":"2.0","id":1,"method":"initialize","params":{
        "protocolVersion":"2024-11-05",
        "capabilities":{},
        "clientInfo":{"name":"claude-helper","version":"1.0"}}}
    code, headers, body = _post(payload)
    # capture session id
    for k,v in headers.items():
        if k.lower() == "mcp-session-id":
            _session_id = v
    # send initialized notification
    _post({"jsonrpc":"2.0","method":"notifications/initialized"})
    return _session_id

# JSON-RPC id 需逐次递增，部分实现会拒绝重复 id
_request_id = {"n": 0}
def _next_id():
    _request_id["n"] += 1
    return _request_id["n"]

def call_tool(name, arguments=None):
    if not _session_id:
        initialize()
    payload = {"jsonrpc":"2.0","id":_next_id(),"method":"tools/call",
               "params":{"name":name,"arguments":arguments or {}}}
    code, headers, body = _post(payload)
    res = _parse(body)
    # tool results come as content array; extract text
    if isinstance(res, dict):
        if "content" in res and isinstance(res["content"], list):
            texts = []
            for c in res["content"]:
                if c.get("type") == "text":
                    texts.append(c.get("text",""))
                elif c.get("type") == "json":
                    texts.append(json.dumps(c.get("json",""), ensure_ascii=False))
            joined = "\n".join(texts)
            # try to parse as JSON for structured results
            try:
                return json.loads(joined)
            except json.JSONDecodeError:
                return joined
        if "error" in res:
            return {"error": res["error"]}
    return res

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: zotero_mcp.py <tool_name> [json_args]")
        sys.exit(1)
    tool = sys.argv[1]
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    sid = initialize()
    print(f"SESSION: {sid}", file=sys.stderr)
    result = call_tool(tool, args)
    print(json.dumps(result, ensure_ascii=False, indent=2) if not isinstance(result, str) else result)
