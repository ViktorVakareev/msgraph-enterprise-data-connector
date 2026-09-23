"""
Step 2 tests — OneDriveFetcher and SharePointFetcher, against a fake
GraphClient session (same _FakeSession/_FakeResponse pattern as
test_graph_client.py). The Excel test uses a real .xlsx file built with
openpyxl and parsed back with pandas — real bytes through the real parsing
path, not a mocked DataFrame.

Run: python test_fetchers.py  (or: pytest -v)
"""

import io

import pandas as pd
import requests

from graph_client import GraphClient
from onedrive_fetcher import OneDriveFetcher
from sharepoint_fetcher import SharePointFetcher


class _FakeAuthenticator:
    def get_auth_headers(self):
        return {"Authorization": "Bearer fake-token"}


class _FakeResponse:
    def __init__(self, status_code=200, json_data=None, content=b""):
        self.status_code = status_code
        self._json_data = json_data
        self.content = content
        self.text = ""

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.exceptions.HTTPError(f"{self.status_code} error")

    def json(self):
        return self._json_data


class _FakeSession:
    """Routes .get() calls by exact URL match rather than a fixed queue —
    fetchers make several distinct calls (list, then content), so exact
    routing is clearer here than the queue used in test_graph_client.py."""

    def __init__(self, routes: dict[str, _FakeResponse]):
        self._routes = routes
        self.calls = []

    def get(self, url, headers=None, params=None, timeout=None):
        self.calls.append({"url": url, "params": params})
        if url not in self._routes:
            raise AssertionError(f"No fake route registered for {url}")
        return self._routes[url]


# ---- OneDriveFetcher ---------------------------------------------------

def _build_sample_xlsx_bytes() -> bytes:
    df = pd.DataFrame({
        "TransactionID": [1, 2, 3],
        "Region": ["West", "East", "West"],
        "Month": ["2026-01", "2026-01", "2026-02"],
        "Sales": [1200.50, 980.00, 1430.25],
    })
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()


def test_onedrive_fetcher_requires_a_user():
    client = GraphClient(_FakeAuthenticator(), session=_FakeSession({}))
    try:
        OneDriveFetcher(client, user="")
    except ValueError as exc:
        assert "app-only auth" in str(exc)
        print(f"[PASS] OneDriveFetcher refuses an empty user (would need /me): {exc}")
    else:
        raise AssertionError("Expected ValueError for empty user")


def test_onedrive_list_files_builds_users_path_not_me():
    list_url = "https://graph.microsoft.com/v1.0/users/agent@contoso.com/drive/root/children"
    session = _FakeSession({list_url: _FakeResponse(200, {"value": [{"id": "f1", "name": "transactions.xlsx"}]})})
    client = GraphClient(_FakeAuthenticator(), session=session)
    fetcher = OneDriveFetcher(client, user="agent@contoso.com")

    files = fetcher.list_files()

    assert files == [{"id": "f1", "name": "transactions.xlsx"}]
    assert "/users/agent@contoso.com/drive/root/children" in session.calls[0]["url"]
    assert "/me/" not in session.calls[0]["url"]
    print("[PASS] list_files() uses /users/{user}/... , never /me/....")


def test_onedrive_fetch_excel_as_dataframe_end_to_end():
    user = "agent@contoso.com"
    list_url = f"https://graph.microsoft.com/v1.0/users/{user}/drive/root/children"
    content_url = f"https://graph.microsoft.com/v1.0/users/{user}/drive/items/f1/content"

    xlsx_bytes = _build_sample_xlsx_bytes()
    session = _FakeSession({
        list_url: _FakeResponse(200, {"value": [{"id": "f1", "name": "transactions.xlsx"}]}),
        content_url: _FakeResponse(200, content=xlsx_bytes),
    })
    client = GraphClient(_FakeAuthenticator(), session=session)
    fetcher = OneDriveFetcher(client, user=user)

    df = fetcher.fetch_excel_as_dataframe("transactions.xlsx")

    assert list(df.columns) == ["TransactionID", "Region", "Month", "Sales"]
    assert len(df) == 3
    assert df.loc[0, "Region"] == "West"
    print(f"[PASS] fetch_excel_as_dataframe() parsed a real .xlsx round trip: {len(df)} rows, columns {list(df.columns)}.")


def test_onedrive_fetch_excel_missing_file_raises_file_not_found():
    list_url = "https://graph.microsoft.com/v1.0/users/agent@contoso.com/drive/root/children"
    session = _FakeSession({list_url: _FakeResponse(200, {"value": []})})
    client = GraphClient(_FakeAuthenticator(), session=session)
    fetcher = OneDriveFetcher(client, user="agent@contoso.com")

    try:
        fetcher.fetch_excel_as_dataframe("does-not-exist.xlsx")
    except FileNotFoundError as exc:
        print(f"[PASS] Missing file correctly raises FileNotFoundError: {exc}")
    else:
        raise AssertionError("Expected FileNotFoundError")


# ---- SharePointFetcher --------------------------------------------------

def test_sharepoint_fetch_list_as_dataframe_end_to_end():
    site_id = "contoso.sharepoint.com,abc-123"
    lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    items_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/list-1/items"

    session = _FakeSession({
        lists_url: _FakeResponse(200, {"value": [{"id": "list-1", "displayName": "RegionalPerformance"}]}),
        items_url: _FakeResponse(200, {"value": [
            {"id": "1", "fields": {"Region": "West", "Month": "2026-01", "TargetSales": 1000, "Status": "On Track"}},
            {"id": "2", "fields": {"Region": "East", "Month": "2026-01", "TargetSales": 900, "Status": "Behind"}},
        ]}),
    })
    client = GraphClient(_FakeAuthenticator(), session=session)
    fetcher = SharePointFetcher(client, site_id=site_id)

    df = fetcher.fetch_list_as_dataframe("RegionalPerformance")

    assert list(df.columns) == ["Region", "Month", "TargetSales", "Status"]
    assert len(df) == 2
    assert session.calls[-1]["params"] == {"$expand": "fields"}
    print(f"[PASS] fetch_list_as_dataframe() built a DataFrame from nested 'fields': {len(df)} rows.")


def test_sharepoint_fetch_list_unknown_name_raises_value_error():
    site_id = "contoso.sharepoint.com,abc-123"
    lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    session = _FakeSession({lists_url: _FakeResponse(200, {"value": []})})
    client = GraphClient(_FakeAuthenticator(), session=session)
    fetcher = SharePointFetcher(client, site_id=site_id)

    try:
        fetcher.fetch_list_as_dataframe("DoesNotExist")
    except ValueError as exc:
        assert "DoesNotExist" in str(exc)
        print(f"[PASS] Unknown list name correctly raises ValueError: {exc}")
    else:
        raise AssertionError("Expected ValueError")


# ---- Step 4: fetcher-level caching --------------------------------------

def test_sharepoint_fetch_list_as_dataframe_cached_avoids_repeat_calls():
    site_id = "contoso.sharepoint.com,abc-123"
    lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    items_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/list-1/items"

    call_log = []

    class _CountingSession(_FakeSession):
        def get(self, url, headers=None, params=None, timeout=None):
            call_log.append(url)
            return super().get(url, headers=headers, params=params, timeout=timeout)

    session = _CountingSession({
        lists_url: _FakeResponse(200, {"value": [{"id": "list-1", "displayName": "RegionalPerformance"}]}),
        items_url: _FakeResponse(200, {"value": [{"id": "1", "fields": {"Region": "West"}}]}),
    })
    client = GraphClient(_FakeAuthenticator(), session=session)
    fetcher = SharePointFetcher(client, site_id=site_id)

    first = fetcher.fetch_list_as_dataframe("RegionalPerformance", use_cache=True)
    second = fetcher.fetch_list_as_dataframe("RegionalPerformance", use_cache=True)

    assert len(first) == len(second) == 1
    # 2 real calls total (lists + items), not 4, even though fetch_list_as_dataframe ran twice
    assert len(call_log) == 2, f"Expected exactly 2 real network calls, got {len(call_log)}: {call_log}"
    print(f"[PASS] fetch_list_as_dataframe(use_cache=True) made {len(call_log)} real calls across 2 invocations.")


if __name__ == "__main__":
    test_onedrive_fetcher_requires_a_user()
    test_onedrive_list_files_builds_users_path_not_me()
    test_onedrive_fetch_excel_as_dataframe_end_to_end()
    test_onedrive_fetch_excel_missing_file_raises_file_not_found()
    test_sharepoint_fetch_list_as_dataframe_end_to_end()
    test_sharepoint_fetch_list_unknown_name_raises_value_error()
    test_sharepoint_fetch_list_as_dataframe_cached_avoids_repeat_calls()
    print("\nAll Step 2 + Step 4 fetcher tests passed.")
