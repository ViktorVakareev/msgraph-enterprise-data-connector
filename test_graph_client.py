"""
Step 2 tests — GraphClient, against a fake `requests`-like session for
precise control over pagination and every failure mode, plus one real
local HTTP server round trip for get_content() (actual bytes over an
actual socket, not just mocked).

Run: python test_graph_client.py  (or: pytest -v)
"""

import http.server
import threading

import requests

from graph_client import GraphClient, GraphRequestError


class _FakeAuthenticator:
    def get_auth_headers(self):
        return {"Authorization": "Bearer fake-token"}


class _FakeResponse:
    def __init__(self, status_code=200, json_data=None, content=b"", raise_for_json_error=False):
        self.status_code = status_code
        self._json_data = json_data
        self.content = content
        self._raise_for_json_error = raise_for_json_error
        self.text = "" if json_data is None else str(json_data)

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.exceptions.HTTPError(f"{self.status_code} error")

    def json(self):
        if self._raise_for_json_error:
            raise ValueError("Expecting value: line 1 column 1 (char 0)")
        return self._json_data


class _FakeSession:
    """Returns a pre-programmed sequence of responses, or raises a
    pre-programmed exception, one per .get() call — in call order."""

    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = []

    def get(self, url, headers=None, params=None, timeout=None):
        self.calls.append({"url": url, "headers": headers, "params": params, "timeout": timeout})
        next_item = self._responses.pop(0)
        if isinstance(next_item, Exception):
            raise next_item
        return next_item


def test_get_success():
    session = _FakeSession([_FakeResponse(200, {"value": [{"id": "1"}]})])
    client = GraphClient(_FakeAuthenticator(), session=session, base_url="https://graph.microsoft.com/v1.0")

    body = client.get("/sites/abc/lists")

    assert body == {"value": [{"id": "1"}]}
    assert session.calls[0]["url"] == "https://graph.microsoft.com/v1.0/sites/abc/lists"
    assert session.calls[0]["headers"] == {"Authorization": "Bearer fake-token"}
    print("[PASS] get() returns parsed JSON and builds the correct URL + auth header.")


def test_get_timeout_raises_graph_request_error():
    session = _FakeSession([requests.exceptions.Timeout("timed out")])
    client = GraphClient(_FakeAuthenticator(), session=session)

    try:
        client.get("/me/messages")
    except GraphRequestError as exc:
        assert "timed out" in str(exc)
        print(f"[PASS] Timeout correctly raised as GraphRequestError: {exc}")
    else:
        raise AssertionError("Expected GraphRequestError")


def test_get_connection_error_raises_graph_request_error():
    session = _FakeSession([requests.exceptions.ConnectionError("refused")])
    client = GraphClient(_FakeAuthenticator(), session=session)

    try:
        client.get("/me/messages")
    except GraphRequestError as exc:
        assert "Could not connect" in str(exc)
        print(f"[PASS] ConnectionError correctly raised as GraphRequestError: {exc}")
    else:
        raise AssertionError("Expected GraphRequestError")


def test_get_http_error_status_raises_graph_request_error():
    session = _FakeSession([_FakeResponse(403, json_data={"error": {"message": "Forbidden"}})])
    client = GraphClient(_FakeAuthenticator(), session=session)

    try:
        client.get("/sites/abc/lists")
    except GraphRequestError as exc:
        assert "403" in str(exc)
        print(f"[PASS] HTTP 403 correctly raised as GraphRequestError: {exc}")
    else:
        raise AssertionError("Expected GraphRequestError")


def test_get_invalid_json_raises_graph_request_error():
    session = _FakeSession([_FakeResponse(200, raise_for_json_error=True)])
    client = GraphClient(_FakeAuthenticator(), session=session)

    try:
        client.get("/sites/abc/lists")
    except GraphRequestError as exc:
        assert "not valid JSON" in str(exc)
        print(f"[PASS] Invalid JSON body correctly raised as GraphRequestError: {exc}")
    else:
        raise AssertionError("Expected GraphRequestError")


def test_get_all_pages_follows_next_link():
    page1 = _FakeResponse(200, {
        "value": [{"id": "1"}, {"id": "2"}],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/abc/lists?$skip=2",
    })
    page2 = _FakeResponse(200, {
        "value": [{"id": "3"}],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/abc/lists?$skip=3",
    })
    page3 = _FakeResponse(200, {"value": [{"id": "4"}]})  # no nextLink: last page

    session = _FakeSession([page1, page2, page3])
    client = GraphClient(_FakeAuthenticator(), session=session)

    items = list(client.get_all_pages("/sites/abc/lists"))

    assert [item["id"] for item in items] == ["1", "2", "3", "4"]
    assert len(session.calls) == 3, "Expected exactly 3 page requests"
    print(f"[PASS] get_all_pages() followed 3 pages and yielded all {len(items)} items in order.")


def test_get_content_over_a_real_local_http_server():
    """get_content() specifically handles raw bytes (file downloads), not
    JSON — verified here over an actual socket, not a mock, since byte
    handling through requests' .content is worth confirming for real."""
    payload = b"PK\x03\x04fake-xlsx-bytes-for-testing"

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.end_headers()
            self.wfile.write(payload)

        def log_message(self, format, *args):
            pass  # silence default request logging

    server = http.server.HTTPServer(("127.0.0.1", 0), Handler)
    port = server.server_port
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    try:
        client = GraphClient(_FakeAuthenticator(), base_url=f"http://127.0.0.1:{port}")
        content = client.get_content("/drive/items/fake-id/content")
        assert content == payload
        print(f"[PASS] get_content() retrieved {len(content)} real bytes over an actual local HTTP server.")
    finally:
        server.shutdown()
        thread.join(timeout=2)


if __name__ == "__main__":
    test_get_success()
    test_get_timeout_raises_graph_request_error()
    test_get_connection_error_raises_graph_request_error()
    test_get_http_error_status_raises_graph_request_error()
    test_get_invalid_json_raises_graph_request_error()
    test_get_all_pages_follows_next_link()
    test_get_content_over_a_real_local_http_server()
    print("\nAll Step 2 GraphClient tests passed.")
