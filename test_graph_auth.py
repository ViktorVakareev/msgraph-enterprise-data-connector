"""
Step 1 tests — verifies GraphAuthenticator against fake MSAL apps, since the
real msal.ConfidentialClientApplication cannot be constructed at all without
live network access to login.microsoftonline.com (confirmed directly; see
the docstring in graph_auth.py). No real Azure tenant is needed to run
these.

Run: python -m pytest test_graph_auth.py -v
"""

from config import GraphConfig
from graph_auth import GraphAuthenticator, GraphAuthError

FAKE_CONFIG = GraphConfig(
    client_id="fake-client-id",
    tenant_id="fake-tenant-id",
    client_secret="fake-secret",
)


class _FakeSuccessApp:
    """Stands in for msal.ConfidentialClientApplication when auth succeeds."""

    def __init__(self, **kwargs):
        self.init_kwargs = kwargs
        self.calls = 0

    def acquire_token_for_client(self, scopes):
        self.calls += 1
        self.last_scopes = scopes
        return {"access_token": "fake-token-value", "expires_in": 3600}


class _FakeFailureApp:
    """Stands in for MSAL returning an error dict (bad secret, missing
    consent, etc.) — MSAL does not raise for this, it returns error keys."""

    def __init__(self, **kwargs):
        pass

    def acquire_token_for_client(self, scopes):
        return {
            "error": "invalid_client",
            "error_description": "AADSTS7000215: Invalid client secret provided.",
        }


def test_get_access_token_success():
    fake_app_instances = []

    def factory(**kwargs):
        app = _FakeSuccessApp(**kwargs)
        fake_app_instances.append(app)
        return app

    authenticator = GraphAuthenticator(FAKE_CONFIG, app_factory=factory)
    token = authenticator.get_access_token()

    assert token == "fake-token-value"
    assert len(fake_app_instances) == 1
    assert fake_app_instances[0].init_kwargs["client_id"] == "fake-client-id"
    assert "fake-tenant-id" in fake_app_instances[0].init_kwargs["authority"]
    print("[PASS] get_access_token returns the token from a successful MSAL call.")


def test_app_is_constructed_only_once():
    """The MSAL app should be built once and reused across calls, not
    reconstructed (which would repeat the expensive tenant-discovery call)
    on every acquire_token_for_client."""
    fake_app_instances = []

    def factory(**kwargs):
        app = _FakeSuccessApp(**kwargs)
        fake_app_instances.append(app)
        return app

    authenticator = GraphAuthenticator(FAKE_CONFIG, app_factory=factory)
    authenticator.get_access_token()
    authenticator.get_access_token()
    authenticator.get_access_token()

    assert len(fake_app_instances) == 1, "MSAL app was reconstructed instead of reused"
    assert fake_app_instances[0].calls == 3
    print("[PASS] The MSAL app instance is built once and reused across calls.")


def test_get_access_token_failure_raises_graph_auth_error():
    authenticator = GraphAuthenticator(FAKE_CONFIG, app_factory=_FakeFailureApp)

    try:
        authenticator.get_access_token()
    except GraphAuthError as exc:
        assert "invalid_client" in str(exc)
        assert "Invalid client secret" in str(exc)
        print(f"[PASS] MSAL error dict correctly raised as GraphAuthError: {exc}")
    else:
        raise AssertionError("Expected GraphAuthError to be raised")


def test_get_auth_headers_format():
    authenticator = GraphAuthenticator(FAKE_CONFIG, app_factory=_FakeSuccessApp)
    headers = authenticator.get_auth_headers()

    assert headers == {"Authorization": "Bearer fake-token-value"}
    print("[PASS] get_auth_headers produces a correctly formatted Bearer header.")


def test_default_scope_is_graph_default():
    fake_app_instances = []

    def factory(**kwargs):
        app = _FakeSuccessApp(**kwargs)
        fake_app_instances.append(app)
        return app

    authenticator = GraphAuthenticator(FAKE_CONFIG, app_factory=factory)
    authenticator.get_access_token()

    assert fake_app_instances[0].last_scopes == ["https://graph.microsoft.com/.default"]
    print("[PASS] Default scope is the Graph .default application-permissions scope.")


if __name__ == "__main__":
    test_get_access_token_success()
    test_app_is_constructed_only_once()
    test_get_access_token_failure_raises_graph_auth_error()
    test_get_auth_headers_format()
    test_default_scope_is_graph_default()
    print("\nAll Step 1 auth tests passed.")
