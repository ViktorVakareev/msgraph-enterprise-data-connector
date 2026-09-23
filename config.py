"""
Step 1 — configuration loading for Graph API access.

Credentials come from environment variables (via a .env file, never
hardcoded), the same pattern used in the Semantic Kernel Foundation
activity's kernel_setup.py: validate everything required up front and raise
one clear error listing exactly what's missing, rather than failing deep
inside an HTTP call with a cryptic 401.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from dotenv import load_dotenv

load_dotenv()

REQUIRED_VARS = ("AZURE_CLIENT_ID", "AZURE_TENANT_ID", "AZURE_CLIENT_SECRET")


@dataclass
class GraphConfig:
    """Credentials and optional resource ids needed to call Microsoft Graph.

    client_id / tenant_id / client_secret come from the Azure Entra ID app
    registration (Implementing-Graph_API.txt walkthrough): the app's
    Application (client) ID, Directory (tenant) ID, and a client secret
    value, generated once and never retrievable again after creation.

    sharepoint_site_id and onedrive_user are optional here — they're only
    needed once Steps 2+ actually call specific SharePoint/OneDrive
    endpoints, not for authentication itself.
    """

    client_id: str
    tenant_id: str
    client_secret: str
    sharepoint_site_id: str | None = None
    onedrive_user: str | None = None


def load_graph_config() -> GraphConfig:
    values = {name: os.environ.get(name) for name in REQUIRED_VARS}
    missing = [name for name, value in values.items() if not value]
    if missing:
        raise RuntimeError(
            "Missing required .env values: "
            + ", ".join(missing)
            + ". Copy .env.example to .env and fill in your Azure Entra ID "
            "app registration's Application (client) ID, Directory (tenant) "
            "ID, and client secret."
        )

    return GraphConfig(
        client_id=values["AZURE_CLIENT_ID"],
        tenant_id=values["AZURE_TENANT_ID"],
        client_secret=values["AZURE_CLIENT_SECRET"],
        sharepoint_site_id=os.environ.get("SHAREPOINT_SITE_ID"),
        onedrive_user=os.environ.get("ONEDRIVE_USER"),
    )


if __name__ == "__main__":
    try:
        config = load_graph_config()
        print(f"Config loaded OK for tenant {config.tenant_id!r}.")
    except RuntimeError as exc:
        print(f"Setup incomplete: {exc}")
