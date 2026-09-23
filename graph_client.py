"""
Step 2/4 — a thin, testable HTTP client for Microsoft Graph.

Step 2 built get()/get_content()/get_all_pages() with unified error handling
(GraphRequestError, chaining the original requests exception). Step 4 adds
the three things this activity's own performance section asks for:

- Timing + logging on every request (surfacing slow calls by default, not
  only when someone goes looking for them).
- Basic backoff-and-retry on HTTP 429 (throttling) — the course's own
  troubleshooting table for this activity names this directly: "Do you see
  HTTP 429 or throttling errors? Add basic backoff (time.sleep() and retry)
  and reduce duplicate calls."
- A simple in-memory cache (get_cached / get_all_pages_cached) for repeated
  identical requests — the same idea as the activity's own
  `@lru_cache`-decorated sample function, implemented here as an explicit
  dict keyed by (path, params) so it's easy to inspect (cache_size()) and
  clear (clear_cache()) in tests, rather than relying on lru_cache's opaque
  per-function cache on a bound method.
"""

from __future__ import annotations

import logging
import time
from typing import Any, Iterator

import requests

logger = logging.getLogger(__name__)

DEFAULT_TIMEOUT_SECONDS = 30
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_MAX_RETRIES = 3
DEFAULT_RETRY_BACKOFF_SECONDS = 1.0


class GraphRequestError(RuntimeError):
    """Raised for any failure calling Microsoft Graph — a timeout, a
    connection error, a non-2xx status (after retries are exhausted for
    429s), or a response body that isn't valid JSON. Always chains the
    original exception via `from exc`."""


class GraphClient:
    """Wraps `GraphAuthenticator` + `requests` behind get()/get_content()/
    get_all_pages(), plus cached and retrying variants, so fetchers
    (OneDriveFetcher, SharePointFetcher) never touch HTTP or auth directly.
    """

    def __init__(
        self,
        authenticator: Any,
        session: Any | None = None,
        base_url: str = GRAPH_BASE_URL,
        timeout: int = DEFAULT_TIMEOUT_SECONDS,
        max_retries: int = DEFAULT_MAX_RETRIES,
        retry_backoff_seconds: float = DEFAULT_RETRY_BACKOFF_SECONDS,
        sleep_fn: Any = time.sleep,
    ) -> None:
        self._authenticator = authenticator
        self._session = session or requests
        self._base_url = base_url.rstrip("/")
        self._timeout = timeout
        self._max_retries = max_retries
        self._retry_backoff_seconds = retry_backoff_seconds
        self._sleep = sleep_fn  # injectable so tests don't actually wait
        self._cache: dict[tuple[str, tuple | None], Any] = {}

    def _resolve_url(self, path_or_url: str) -> str:
        # @odata.nextLink values are already full URLs; a plain resource
        # path (e.g. "/sites/{id}/lists") needs the base prepended.
        if path_or_url.startswith("http://") or path_or_url.startswith("https://"):
            return path_or_url
        return f"{self._base_url}/{path_or_url.lstrip('/')}"

    @staticmethod
    def _parse_retry_after(response: Any) -> float | None:
        value = getattr(response, "headers", {}).get("Retry-After") if hasattr(response, "headers") else None
        if value is None:
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    def _request(self, url: str, params: dict[str, Any] | None = None) -> Any:
        """Performs the GET, with logging + timing on every attempt, and
        backoff-and-retry on HTTP 429 up to `max_retries`. Returns the raw
        response object; callers (get/get_content) decide how to read it."""
        headers = self._authenticator.get_auth_headers()
        attempt = 0

        while True:
            start = time.perf_counter()
            try:
                response = self._session.get(url, headers=headers, params=params, timeout=self._timeout)
            except requests.exceptions.Timeout as exc:
                raise GraphRequestError(f"Graph request to {url} timed out") from exc
            except requests.exceptions.ConnectionError as exc:
                raise GraphRequestError(f"Could not connect to {url}") from exc
            except requests.exceptions.RequestException as exc:
                raise GraphRequestError(f"Graph request to {url} failed") from exc

            elapsed = time.perf_counter() - start
            logger.info("Graph GET %s -> %s in %.3fs", url, response.status_code, elapsed)

            if response.status_code == 429 and attempt < self._max_retries:
                delay = self._parse_retry_after(response) or (self._retry_backoff_seconds * (2 ** attempt))
                logger.warning(
                    "Graph throttled (429) on %s — retrying in %.1fs (attempt %d/%d)",
                    url, delay, attempt + 1, self._max_retries,
                )
                self._sleep(delay)
                attempt += 1
                continue

            try:
                response.raise_for_status()
            except requests.exceptions.HTTPError as exc:
                raise GraphRequestError(
                    f"Graph request to {url} failed with status {response.status_code}: {response.text[:300]}"
                ) from exc

            return response

    def get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        """A single GET returning the parsed JSON body."""
        url = self._resolve_url(path)
        response = self._request(url, params=params)
        try:
            return response.json()
        except ValueError as exc:
            raise GraphRequestError(f"Graph response from {url} was not valid JSON") from exc

    def get_content(self, path: str) -> bytes:
        """A GET against a /content endpoint, returning raw bytes (for file
        downloads) instead of parsed JSON."""
        url = self._resolve_url(path)
        response = self._request(url)
        return response.content

    def get_all_pages(self, path: str, params: dict[str, Any] | None = None) -> Iterator[dict[str, Any]]:
        """Yields every item across all pages of a Graph collection
        response, following "@odata.nextLink" until it's absent. Graph
        paginates any collection past its page-size limit (commonly 200
        items) — a fetcher that only reads `.get("value", [])` once would
        silently truncate large SharePoint lists or OneDrive folders.

        Known bottleneck (documented, not yet fixed here — out of this
        activity's scope): pages are fetched serially, one full network
        round trip at a time. For a very large list this is the single
        biggest cost in the whole pipeline; a production version would
        either request a larger page size (`$top`) or fetch pages
        concurrently.
        """
        next_path: str | None = path
        next_params: dict[str, Any] | None = params

        while next_path is not None:
            body = self.get(next_path, params=next_params)
            for item in body.get("value", []):
                yield item

            next_path = body.get("@odata.nextLink")
            next_params = None  # nextLink is already a complete URL with its own query string

    # ---- Caching (Step 4) --------------------------------------------------
    #
    # Known bottleneck (documented, and fixed by this cache): a pipeline that
    # calls the same endpoint with the same parameters more than once — e.g.
    # re-listing a OneDrive folder to look up two different files by name —
    # pays a full network round trip each time. Caching the materialized
    # result means only the first call ever touches the network.

    @staticmethod
    def _cache_key(path: str, params: dict[str, Any] | None) -> tuple[str, tuple | None]:
        params_key = tuple(sorted(params.items())) if params else None
        return (path, params_key)

    def get_cached(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        key = self._cache_key(path, params)
        if key in self._cache:
            logger.debug("Cache hit for %s", key)
            return self._cache[key]
        result = self.get(path, params=params)
        self._cache[key] = result
        return result

    def get_all_pages_cached(self, path: str, params: dict[str, Any] | None = None) -> list[dict[str, Any]]:
        key = self._cache_key(path, params)
        if key in self._cache:
            logger.debug("Cache hit for %s", key)
            return list(self._cache[key])
        items = list(self.get_all_pages(path, params=params))
        self._cache[key] = items
        return list(items)

    def clear_cache(self) -> None:
        self._cache.clear()

    def cache_size(self) -> int:
        return len(self._cache)
