import os
import sys
import unittest
from fastapi.testclient import TestClient

# Add parent directory
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from app import app
from wcf_integrator import WCFIntegrator

class TestAppAPI(unittest.TestCase):

    def setUp(self):
        self.client = TestClient(app)

    def test_serve_index(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn("SSO Debt Report System", response.text)

    def test_get_system_info(self):
        response = self.client.get("/api/system-info")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["wcf_connected"])
        expected = WCFIntegrator().total_records
        self.assertEqual(data["wcf_records_count"], expected)
        self.assertGreater(expected, 0)
        self.assertIsInstance(data["input_files"], list)
        filenames = [f["filename"] for f in data["input_files"]]
        self.assertIn("CN200090-1_20260914_161817.xlsx", filenames)

    def test_process_existing_file(self):
        response = self.client.post(
            "/api/process-existing",
            json={"filename": "CN200090-1_20260914_161817.xlsx"}
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertEqual(data["total_establishments"], 152)
        self.assertEqual(data["matched_wcf_count"], 148)
        self.assertAlmostEqual(data["total_principal_debt"], 1885246.00, places=2)

        # Test download
        download_res = self.client.get(f"/api/download/{data['output_filename']}")
        self.assertEqual(download_res.status_code, 200)
        self.assertEqual(
            download_res.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Cleanup test output file
        from app import OUTPUT_DIR
        out_f = os.path.join(OUTPUT_DIR, data['output_filename'])
        if os.path.exists(out_f):
            os.remove(out_f)

    def test_reload_master_data(self):
        response = self.client.post("/api/reload-master-data")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        expected = WCFIntegrator().total_records
        self.assertEqual(data["wcf_records_count"], expected)
        self.assertIn(f"{expected:,}", data["message"])

    def test_upload_master_data_validation(self):
        # Invalid file type
        response = self.client.post(
            "/api/upload-master-data",
            files={"file": ("test.txt", b"invalid data", "text/plain")}
        )
        self.assertEqual(response.status_code, 400)

    def test_path_traversal_protection(self):
        # Path traversal download attempt
        response = self.client.get("/api/download/../../etc/passwd")
        self.assertIn(response.status_code, (404, 400))

        # Path traversal process attempt
        response2 = self.client.post(
            "/api/process-existing",
            json={"filename": "../../app.py"}
        )
        self.assertEqual(response2.status_code, 404)

    def test_letter_config_endpoints(self):
        # 1. GET config
        res = self.client.get("/api/letter-config")
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertEqual(data["status"], "success")
        self.assertIn("config", data)
        self.assertIn("doc_code", data["config"])
        self.assertIn("office_name", data["config"])

        # 2. POST update config
        update_payload = {
            "doc_code": "สน ๐๐๓1.1/สท-ทดสอบ",
            "signatory_name": "นายชลาลัย ประโมสี (ทดสอบ)",
        }
        res_post = self.client.post("/api/letter-config", json=update_payload)
        self.assertEqual(res_post.status_code, 200)
        post_data = res_post.json()
        self.assertEqual(post_data["config"]["doc_code"], "สน ๐๐๓1.1/สท-ทดสอบ")
        self.assertEqual(post_data["config"]["signatory_name"], "นายชลาลัย ประโมสี (ทดสอบ)")

        # Restore default
        self.client.post("/api/letter-config", json={
            "doc_code": "สน ๐๐๓1.1/สท",
            "signatory_name": "นายชลาลัย ประโมสี",
        })

if __name__ == "__main__":
    unittest.main()
