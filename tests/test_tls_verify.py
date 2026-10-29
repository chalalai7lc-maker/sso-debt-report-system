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
