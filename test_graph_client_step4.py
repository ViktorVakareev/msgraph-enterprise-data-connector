"""
Step 4 tests — GraphClient's retry/backoff on 429 throttling, and the
in-memory cache. Uses the same _FakeSession/_FakeResponse pattern as
test_graph_client.py, with a fake `sleep_fn` injected so these tests run
instantly instead of actually waiting through the backoff delays.

Run: python test_graph_client_step4.py  (or: pytest -v)
"""

import requests

from graph_client import GraphClient, GraphRequestError


class _FakeAuthenticator:
    def get_auth_headers(self):
        return {"Authorization": "Bearer fake-token"}


class _FakeResponse:
    def __init__(self, status_code=200, json_data=None, headers=None):
        self.status_code = status_code
        self._json_data = json_data
        self.headers = headers or {}
        self.text = "" if json_data is None else str(json_data)
        self.content = b""

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.exceptions.HTTPError(f"{self.status_code} error")

    def json(self):
        return self._json_data


class _FakeSession:
    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = []

    def get(self, url, headers=None, params=None, timeout=None):
        self.calls.append({"url": url, "params": params})
        return self._responses.pop(0)


class _FakeSleep:
    """Records requested delays instead of actually sleeping."""

    def __init__(self):
        self.delays = []

    def __call__(self, seconds):
        self.delays.append(seconds)


# ---- Retry / backoff on 429 --------------------------------------------

def test_429_retries_then_succeeds():
    session = _FakeSession([
        _FakeResponse(429, headers={}),
        _FakeResponse(429, headers={}),
        _FakeResponse(200, {"value": [{"id": "1"}]}),
    ])
    sleep = _FakeSleep()
    client = GraphClient(_FakeAuthenticator(), session=session, sleep_fn=sleep, retry_backoff_seconds=0.5)

    body = client.get("/sites/abc/lists")

    assert body == {"value": [{"id": "1"}]}
    assert len(session.calls) == 3, "Expected 2 throttled attempts + 1 success"
    # exponential backoff: 0.5, then 1.0 (doubling each retry)
    assert sleep.delays == [0.5, 1.0]
    print(f"[PASS] 429 responses triggered exactly 2 retries with exponential backoff {sleep.delays}, then succeeded.")


def test_429_honors_retry_after_header():
    session = _FakeSession([
        _FakeResponse(429, headers={"Retry-After": "3"}),
        _FakeResponse(200, {"value": []}),
    ])
    sleep = _FakeSleep()
    client = GraphClient(_FakeAuthenticator(), session=session, sleep_fn=sleep)

    client.get("/sites/abc/lists")

    assert sleep.delays == [3.0]
    print(f"[PASS] Retry-After header value (3s) is honored instead of the default backoff: {sleep.delays}")


def test_429_gives_up_after_max_retries():
    # max_retries=2 means at most 2 retries: 3 total 429 responses exhausts it
    session = _FakeSession([
        _FakeResponse(429, headers={}),
        _FakeResponse(429, headers={}),
        _FakeResponse(429, headers={}),
    ])
    sleep = _FakeSleep()
    client = GraphClient(_FakeAuthenticator(), session=session, sleep_fn=sleep, max_retries=2, retry_backoff_seconds=0.1)

    try:
        client.get("/sites/abc/lists")
    except GraphRequestError as exc:
        assert "429" in str(exc)
        assert len(session.calls) == 3  # initial + 2 retries, then it gives up
        print(f"[PASS] After max_retries is exhausted, the 429 is correctly raised as GraphRequestError: {exc}")
    else:
        raise AssertionError("Expected GraphRequestError after exhausting retries")


# ---- Caching ------------------------------------------------------------

def test_get_cached_avoids_repeat_network_calls():
    session = _FakeSession([_FakeResponse(200, {"value": [{"id": "1"}]})])
    client = GraphClient(_FakeAuthenticator(), session=session)

    first = client.get_cached("/sites/abc/lists")
    second = client.get_cached("/sites/abc/lists")
    third = client.get_cached("/sites/abc/lists")

    assert first == second == third == {"value": [{"id": "1"}]}
    assert len(session.calls) == 1, "Expected exactly 1 real network call; the rest should be served from cache"
    assert client.cache_size() == 1
    print(f"[PASS] get_cached() made 1 real network call across 3 identical requests (cache_size={client.cache_size()}).")


def test_get_cached_different_params_are_separate_entries():
    session = _FakeSession([
        _FakeResponse(200, {"value": [{"id": "west"}]}),
        _FakeResponse(200, {"value": [{"id": "east"}]}),
    ])
    client = GraphClient(_FakeAuthenticator(), session=session)

    west = client.get_cached("/sites/abc/lists", params={"$filter": "region eq 'West'"})
    east = client.get_cached("/sites/abc/lists", params={"$filter": "region eq 'East'"})

    assert west != east
    assert len(session.calls) == 2
    assert client.cache_size() == 2
    print("[PASS] Different params produce separate cache entries rather than colliding.")


def test_get_all_pages_cached_avoids_repeat_pagination():
    page1 = _FakeResponse(200, {"value": [{"id": "1"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"})
    page2 = _FakeResponse(200, {"value": [{"id": "2"}]})
    session = _FakeSession([page1, page2])
    client = GraphClient(_FakeAuthenticator(), session=session)

    first_call = client.get_all_pages_cached("/sites/abc/lists")
    second_call = client.get_all_pages_cached("/sites/abc/lists")

    assert [item["id"] for item in first_call] == ["1", "2"]
    assert first_call == second_call
    assert len(session.calls) == 2, "Expected exactly 2 real page requests on the first call only"
    print("[PASS] get_all_pages_cached() paginated fully once, then served the second call entirely from cache.")


def test_clear_cache_forces_a_fresh_network_call():
    session = _FakeSession([
        _FakeResponse(200, {"value": [{"id": "1"}]}),
        _FakeResponse(200, {"value": [{"id": "1-updated"}]}),
    ])
    client = GraphClient(_FakeAuthenticator(), session=session)

    client.get_cached("/sites/abc/lists")
    client.clear_cache()
    result = client.get_cached("/sites/abc/lists")

    assert result == {"value": [{"id": "1-updated"}]}
    assert len(session.calls) == 2
    print("[PASS] clear_cache() correctly forces the next call back out to the network.")


if __name__ == "__main__":
    test_429_retries_then_succeeds()
    test_429_honors_retry_after_header()
    test_429_gives_up_after_max_retries()
    test_get_cached_avoids_repeat_network_calls()
    test_get_cached_different_params_are_separate_entries()
    test_get_all_pages_cached_avoids_repeat_pagination()
    test_clear_cache_forces_a_fresh_network_call()
    print("\nAll Step 4 GraphClient tests passed.")
