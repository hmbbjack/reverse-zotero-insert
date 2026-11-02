#!/usr/bin/env python3
"""Streamable HTTP MCP client with explicit transport and protocol errors."""

import argparse
import itertools
import json
import sys
import urllib.error
import urllib.request

if __package__:
    from . import config
else:
    import config


class MCPError(RuntimeError):
    """The server could not complete an MCP request."""


def _result(message, request_id):
    if not isinstance(message, dict) or message.get("id") != request_id:
        raise MCPError("MCP response does not match the request ID")
    if "error" in message:
        raise MCPError(f"MCP error: {message['error']}")
    if "result" not in message:
        raise MCPError("MCP response has no result")
    return message["result"]


def _read_sse(response, request_id):
    """Stop at the matching event without waiting for the server to close the stream."""
    data = []
    for raw_line in response:
        line = raw_line.decode("utf-8").rstrip("\r\n")
        if line.startswith("data:"):
            data.append(line[5:].lstrip(" "))
        elif not line and data:
            message = json.loads("\n".join(data))
            data = []
            if isinstance(message, dict) and message.get("id") == request_id:
                return _result(message, request_id)
    if data:
        return _result(json.loads("\n".join(data)), request_id)
    raise MCPError("MCP stream ended without a matching response")


class MCPClient:
    """One sequential session; stateless servers may omit a session ID."""

    def __init__(self, url=None, timeout=60):
        self.url = url or config.get_mcp_url()
        self.timeout = timeout
        self.session_id = None
        self.protocol_version = "2024-11-05"
        self._ids = itertools.count(1)
        self._initialized = False

    def _post(self, method, params=None, *, notification=False):
        payload = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            payload["params"] = params
        if not notification:
            payload["id"] = next(self._ids)
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
        }
        if self.session_id:
            headers["Mcp-Session-Id"] = self.session_id
        if self._initialized or method == "notifications/initialized":
            headers["MCP-Protocol-Version"] = self.protocol_version
        request = urllib.request.Request(
            self.url, json.dumps(payload).encode(), headers, method="POST"
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                if notification:
                    return None
                if method == "initialize":
                    self.session_id = response.headers.get("Mcp-Session-Id")
                content_type = response.headers.get("Content-Type", "")
                if "text/event-stream" in content_type:
                    return _read_sse(response, payload["id"])
                return _result(json.load(response), payload["id"])
        except urllib.error.HTTPError as error:
            error.close()
            raise MCPError(f"MCP HTTP {error.code} during {method}") from error
        except (urllib.error.URLError, OSError, ValueError) as error:
            raise MCPError(f"MCP {method} failed: {error}") from error

    def initialize(self):
        if not self._initialized:
            result = self._post(
                "initialize",
                {
                    "protocolVersion": self.protocol_version,
                    "capabilities": {},
                    "clientInfo": {"name": "reverse-zotero-insert", "version": "3.1.0"},
                },
            )
            if not isinstance(result, dict) or not result.get("protocolVersion"):
                raise MCPError("Invalid MCP initialize response")
            self.protocol_version = result["protocolVersion"]
            self._post("notifications/initialized", notification=True)
            self._initialized = True
        return self.session_id

    def call_tool(self, name, arguments=None):
        self.initialize()
        result = self._post("tools/call", {"name": name, "arguments": arguments or {}})
        if not isinstance(result, dict):
            raise MCPError("Invalid MCP tool response")
        if result.get("isError"):
            raise MCPError(f"MCP tool {name} failed: {result.get('content', [])}")
        if "structuredContent" in result:
            decoded = result["structuredContent"]
        elif "content" in result:
            texts = []
            for block in result["content"]:
                if block.get("type") == "text":
                    texts.append(block.get("text", ""))
                elif block.get("type") == "json":
                    texts.append(json.dumps(block.get("json"), ensure_ascii=False))
            text = "\n".join(texts)
            try:
                decoded = json.loads(text)
            except json.JSONDecodeError:
                return text
        else:
            decoded = result
        if isinstance(decoded, dict) and decoded.get("error"):
            raise MCPError(f"MCP tool {name} failed: {decoded['error']}")
        return decoded


_default_client = None


def _client():
    global _default_client
    if _default_client is None:
        _default_client = MCPClient()
    return _default_client


def initialize():
    return _client().initialize()


def call_tool(name, arguments=None):
    return _client().call_tool(name, arguments)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("tool")
    parser.add_argument("arguments", nargs="?", default="{}", type=json.loads)
    args = parser.parse_args(argv)
    try:
        result = call_tool(args.tool, args.arguments)
    except MCPError as error:
        print(str(error), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
