"""
Step 2 — OneDrive file listing and Excel retrieval via Microsoft Graph.

Uses /users/{user}/drive/... rather than /me/drive/.... This project's auth
(Step 1) is the client-credentials, app-only flow — there is no signed-in
user — and the course's own troubleshooting table for this exact activity
calls this out directly: "/me not supported ... Using /me endpoint without
appropriate permissions -> Switch to /users/". OneDriveFetcher enforces
this by requiring a `user` (UPN or object id) at construction, rather than
letting a caller accidentally build a /me/... path that will 410 in
production.
"""

from __future__ import annotations

import io

import pandas as pd

from graph_client import GraphClient


class OneDriveFetcher:
    def __init__(self, client: GraphClient, user: str) -> None:
        if not user:
            raise ValueError(
                "OneDriveFetcher requires a user (UPN or object id) for app-only auth — "
                "the /me endpoint is not usable without a signed-in user."
            )
        self._client = client
        self._user = user

    def _drive_root(self) -> str:
        return f"/users/{self._user}/drive/root"

    def list_files(self, folder_path: str | None = None, use_cache: bool = False) -> list[dict]:
        """Lists files (and subfolders) directly inside a OneDrive folder.
        `folder_path` is a path relative to the drive root, e.g.
        "Reports/2026"; omit it to list the root folder.

        `use_cache=True` (Step 4) avoids repeating this call's full
        pagination if it's made again with the same folder_path — useful
        since `get_file_metadata_by_name` calls this once per lookup, and a
        pipeline commonly looks up several files from the same folder."""
        if folder_path:
            path = f"{self._drive_root()}:/{folder_path.strip('/')}:/children"
        else:
            path = f"{self._drive_root()}/children"

        if use_cache:
            return self._client.get_all_pages_cached(path)
        return list(self._client.get_all_pages(path))

    def get_file_metadata_by_name(
        self, filename: str, folder_path: str | None = None, use_cache: bool = False
    ) -> dict | None:
        for item in self.list_files(folder_path, use_cache=use_cache):
            if item.get("name") == filename:
                return item
        return None

    def download_file_content(self, item_id: str) -> bytes:
        return self._client.get_content(f"/users/{self._user}/drive/items/{item_id}/content")

    def fetch_excel_as_dataframe(
        self,
        filename: str,
        folder_path: str | None = None,
        sheet_name: str | int = 0,
        use_cache: bool = False,
    ) -> pd.DataFrame:
        """Finds `filename` in the given OneDrive folder, downloads its raw
        bytes, and parses it as an Excel workbook — no temp file on disk,
        matching how a real Graph-backed pipeline would run in a
        stateless/containerized job."""
        metadata = self.get_file_metadata_by_name(filename, folder_path, use_cache=use_cache)
        if metadata is None:
            location = f" in folder {folder_path!r}" if folder_path else ""
            raise FileNotFoundError(f"No file named {filename!r} found{location} in this OneDrive.")

        content = self.download_file_content(metadata["id"])
        return pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, engine="openpyxl")
