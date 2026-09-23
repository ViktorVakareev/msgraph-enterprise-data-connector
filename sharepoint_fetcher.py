"""
Step 2 — SharePoint list enumeration and item retrieval via Microsoft Graph.

GET /sites/{site-id}/lists enumerates the lists on a site; GET
/sites/{site-id}/lists/{list-id}/items?$expand=fields returns each item's
actual column data under a nested "fields" object — Graph keeps this
metadata separate from field values by design, so `$expand=fields` is not
optional if the goal is the list's real content rather than bare item ids.
"""

from __future__ import annotations

import pandas as pd

from graph_client import GraphClient


class SharePointFetcher:
    def __init__(self, client: GraphClient, site_id: str) -> None:
        if not site_id:
            raise ValueError("SharePointFetcher requires a site_id.")
        self._client = client
        self._site_id = site_id

    def list_lists(self, use_cache: bool = False) -> list[dict]:
        path = f"/sites/{self._site_id}/lists"
        if use_cache:
            return self._client.get_all_pages_cached(path)
        return list(self._client.get_all_pages(path))

    def get_list_id_by_name(self, display_name: str, use_cache: bool = False) -> str | None:
        for lst in self.list_lists(use_cache=use_cache):
            if lst.get("displayName") == display_name or lst.get("name") == display_name:
                return lst.get("id")
        return None

    def list_items(self, list_id: str, expand_fields: bool = True, use_cache: bool = False) -> list[dict]:
        params = {"$expand": "fields"} if expand_fields else None
        path = f"/sites/{self._site_id}/lists/{list_id}/items"
        if use_cache:
            return self._client.get_all_pages_cached(path, params=params)
        return list(self._client.get_all_pages(path, params=params))

    def fetch_list_as_dataframe(self, list_name: str, use_cache: bool = False) -> pd.DataFrame:
        """Looks up a list by its display name and returns its items' field
        data as a DataFrame — one row per item, one column per field.

        `use_cache=True` (Step 4) avoids re-fetching both the site's list
        catalog and the item pagination if this is called more than once
        for the same list — e.g. re-running analysis against data already
        fetched earlier in the same process."""
        list_id = self.get_list_id_by_name(list_name, use_cache=use_cache)
        if list_id is None:
            raise ValueError(f"No SharePoint list named {list_name!r} found on site {self._site_id!r}.")

        items = self.list_items(list_id, use_cache=use_cache)
        rows = [item.get("fields", {}) for item in items]
        return pd.DataFrame(rows)
