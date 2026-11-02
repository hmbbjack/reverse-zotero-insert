"""Exercise MCP wire responses without connecting to a real library."""

import io
import json
import unittest
from unittest.mock import patch
from urllib.error import HTTPError

import zotero_mcp as mcp


class Response(io.BytesIO):
    def __init__(self, payload, headers=None):
        super().__init__(json.dumps(payload).encode())
        self.headers = {"Content-Type": "application/json", **(headers or {})}


class ClientTests(unittest.TestCase):
    def server(self, session=None, tool_result=None):
        requests = []

        def open_request(request, **kwargs):
            payload = json.loads(request.data)
            requests.append(payload)
            if payload["method"] == "initialize":
                return Response(
                    {"id": payload["id"], "result": {"protocolVersion": "2024-11-05"}},
                    {"Mcp-Session-Id": session} if session else {},
                )
            if "id" not in payload:
                return Response({})
            result = (
                tool_result
                if tool_result is not None
                else {"content": [{"type": "text", "text": '{"items": []}'}]}
            )
            return Response({"id": payload["id"], "result": result})

        return requests, open_request

    def test_unique_ids_and_single_initialization_with_or_without_session(self):
        for session in (None, "test-session"):
            with self.subTest(session=session):
                requests, service = self.server(session)
                with patch.object(mcp.urllib.request, "urlopen", service):
                    client = mcp.MCPClient()
                    self.assertEqual(client.call_tool("get_collections"), {"items": []})
                    client.call_tool("get_collections")
                self.assertEqual([r["id"] for r in requests if "id" in r], [1, 2, 3])
                self.assertEqual(sum(r["method"] == "initialize" for r in requests), 1)

    def test_failed_initialization_never_calls_tool(self):
        with patch.object(
            mcp.urllib.request,
            "urlopen",
            side_effect=HTTPError("http://localhost/mcp", 503, "unavailable", {}, io.BytesIO()),
        ) as transport:
            with self.assertRaisesRegex(mcp.MCPError, "HTTP 503"):
                mcp.MCPClient().call_tool("write_item")
        self.assertEqual(transport.call_count, 1)

    def test_tool_errors_are_not_returned_as_successful_text(self):
        for result in (
            {"isError": True, "content": [{"type": "text", "text": "failure"}]},
            {"structuredContent": {"error": "unavailable"}},
        ):
            with self.subTest(result=result):
                _, service = self.server(tool_result=result)
                with patch.object(mcp.urllib.request, "urlopen", service):
                    with self.assertRaises(mcp.MCPError):
                        mcp.MCPClient().call_tool("get_collections")

    def test_structured_content_is_supported(self):
        _, service = self.server(tool_result={"structuredContent": {"items": []}})
        with patch.object(mcp.urllib.request, "urlopen", service):
            self.assertEqual(mcp.MCPClient().call_tool("get_collections"), {"items": []})

    def test_mismatched_response_id_is_rejected(self):
        with self.assertRaises(mcp.MCPError):
            mcp._result({"id": 5, "result": {}}, 6)

    def test_sse_ignores_notifications_and_stops_before_connection_close(self):
        def stream():
            yield b": keepalive\n"
            yield b'data: {"jsonrpc":"2.0","method":"notifications/progress"}\n'
            yield b"\n"
            yield b"event: message\n"
            yield b'data: {"id":2,\n'
            yield b'data: "result":{"items":[]}}\n'
            yield b"\n"
            raise AssertionError("Client waited for an SSE connection to close")

        self.assertEqual(mcp._read_sse(stream(), 2), {"items": []})

    def test_json_rpc_error_is_raised(self):
        with self.assertRaisesRegex(mcp.MCPError, "unknown tool"):
            mcp._result({"id": 2, "error": {"message": "unknown tool"}}, 2)


if __name__ == "__main__":
    unittest.main()
