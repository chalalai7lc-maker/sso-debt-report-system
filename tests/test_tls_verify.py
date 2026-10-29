import inspect, os, unittest
from unittest.mock import patch
import sso_client

class TLSVerifyTest(unittest.TestCase):
    def test_no_verify_false_in_source(self):
        self.assertNotIn("verify=False", inspect.getsource(sso_client))
    def test_default_verify_on(self):
        with patch.dict(os.environ, {"SSO_CA_BUNDLE": ""}):
            self.assertIs(sso_client.tls_verify_setting(), True)
    def test_ca_bundle_missing_fails_closed(self):
        with patch.dict(os.environ, {"SSO_CA_BUNDLE": "/nonexistent/ca.pem"}):
            with self.assertRaises(RuntimeError):
                sso_client.tls_verify_setting()
    def test_requirements_no_httpx2(self):
        req = open(os.path.join(os.path.dirname(__file__), "..", "requirements.txt")).read().split()
        self.assertFalse(any(l.startswith("httpx2") for l in req))


# --- round 3: Playwright + startup CA validation ---
import ssl as _ssl
import pytest as _pytest
import sso_client as _sc


def test_playwright_never_ignores_tls(monkeypatch):
    monkeypatch.delenv("SSO_CA_BUNDLE", raising=False)
    assert _sc.playwright_ignore_https_errors() is False


def test_playwright_with_valid_bundle_still_verifies(monkeypatch):
    import certifi
    monkeypatch.setenv("SSO_CA_BUNDLE", certifi.where())
    assert _sc.playwright_ignore_https_errors() is False
    assert _sc.tls_verify_setting() == certifi.where()


def test_invalid_pem_bundle_rejected(monkeypatch, tmp_path):
    bad = tmp_path / "bad.pem"
    bad.write_text("not a certificate")
    monkeypatch.setenv("SSO_CA_BUNDLE", str(bad))
    with _pytest.raises(RuntimeError):
        _sc.tls_verify_setting()
    with _pytest.raises(RuntimeError):
        _sc.playwright_ignore_https_errors()
    with _pytest.raises(RuntimeError):
        _sc.validate_tls_config()


def test_missing_bundle_fails_startup(monkeypatch):
    monkeypatch.setenv("SSO_CA_BUNDLE", "/nonexistent/ca.pem")
    with _pytest.raises(RuntimeError):
        _sc.validate_tls_config()


def test_source_has_no_hardcoded_ignore():
    import inspect
    assert "ignore_https_errors=True" not in inspect.getsource(_sc)
