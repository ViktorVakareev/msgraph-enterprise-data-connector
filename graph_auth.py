"""
Step 1 — Microsoft Graph authentication, via MSAL's client-credentials
(app-only) flow.

The Entra ID walkthrough (Implementing-Graph_API.txt) specifically
registers this agent with **Application permissions** — "for an automated
agent that runs without a user present" — rather than delegated
permissions. That's the client-credentials grant: the app authenticates as
itself (client ID + client secret), not on behalf of a signed-in user, and
MSAL's ConfidentialClientApplication.acquire_token_for_client(...) is the
matching call.

Testing note (a real finding, not a guess): constructing
msal.ConfidentialClientApplication performs a live HTTPS call to
login.microsoftonline.com immediately (tenant discovery), even before any
token request. Confirmed directly: it fails in a network-restricted
environment with `ProxyError` regardless of `validate_authority`. That
means unit tests can't just mock `requests` underneath MSAL — they have to
inject a fake in place of the whole `ConfidentialClientApplication` class.
That's why `app_factory` is a constructor parameter here rather than a
hardcoded import-time call: production code passes nothing (defaults to
the real MSAL class), tests pass a stub.
"""

from __future__ import annotations

from typing import Any, Callable

import msal

from config import GraphConfig

GRAPH_DEFAULT_SCOPE = ["https://graph.microsoft.com/.default"]
AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}"


class GraphAuthError(RuntimeError):
    """Raised when MSAL returns an error instead of an access token.

    MSAL does not raise Python exceptions for auth failures (bad secret,
    missing admin consent, etc.) — it returns a dict containing "error" and
    "error_description" keys instead. Converting that into a real exception
    here means callers only ever have to handle one failure path.
    """


class GraphAuthenticator:
    """Wraps MSAL's client-credentials flow behind a small, testable interface."""

    def __init__(
        self,
        config: GraphConfig,
        scopes: list[str] | None = None,
        app_factory: Callable[..., Any] | None = None,
    ) -> None:
        self._config = config
        self._scopes = scopes or GRAPH_DEFAULT_SCOPE
        self._app_factory = app_factory or msal.ConfidentialClientApplication
        self._app: Any = None

    def _get_app(self) -> Any:
        if self._app is None:
            self._app = self._app_factory(
                client_id=self._config.client_id,
                client_credential=self._config.client_secret,
                authority=AUTHORITY_TEMPLATE.format(tenant_id=self._config.tenant_id),
            )
        return self._app

    def get_access_token(self) -> str:
        """Returns a bearer access token, acquiring (or reusing MSAL's
        internal cache of) one via the client-credentials flow."""
        app = self._get_app()
        result = app.acquire_token_for_client(scopes=self._scopes)

        if "access_token" not in result:
            error = result.get("error", "unknown_error")
            description = result.get("error_description", "No further details returned.")
            raise GraphAuthError(f"Failed to acquire Graph access token ({error}): {description}")

        return result["access_token"]

    def get_auth_headers(self) -> dict[str, str]:
        """The Authorization header every Graph API call needs."""
        return {"Authorization": f"Bearer {self.get_access_token()}"}


if __name__ == "__main__":
    from config import load_graph_config

    try:
        cfg = load_graph_config()
    except RuntimeError as exc:
        print(f"Setup incomplete: {exc}")
    else:
        authenticator = GraphAuthenticator(cfg)
        try:
            token = authenticator.get_access_token()
            print(f"Access token acquired ({len(token)} chars). Auth is working.")
        except GraphAuthError as exc:
            print(f"Authentication failed: {exc}")
