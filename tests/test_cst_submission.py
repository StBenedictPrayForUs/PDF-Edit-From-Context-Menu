import json
import tempfile
import threading
import unittest
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

from app.cst_submission import discard_batch, submit_batch
from app.cst_workflow import export_cst_batch
from tests.test_cst_workflow import make_pdf


class FakeApi:
    """Mimics the server: the first payload for a submission ID wins."""
    def __init__(self):
        self.receipts = {}
        self.posts = []
        self.lose_response = False

    def post(self, payload):
        self.posts.append(payload)
        submission_id = payload["submissionId"]
        if submission_id in self.receipts:
            return 200, {**self.receipts[submission_id], "duplicate": True}
        receipt = {"accepted": True, "submissionId": submission_id,
                   "savedToServer": True, "status": "received"}
        self.receipts[submission_id] = receipt
        if self.lose_response:
            self.lose_response = False
            raise TimeoutError("lost response")
        return 201, receipt


class CstSubmissionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        source = root / "source.pdf"
        make_pdf(source)
        self.batch = export_cst_batch(source, None, date(2026, 9, 17),
            [(1, 2, "Alpha / Exact", "Tracking"), (3, 4, "Beta", "Delivery")], ["Alpha / Exact", "Beta"],
            [1, 2, 3, 4], {}, root / "out", technician="Ryan") / "batch.json"

    def tearDown(self):
        self.temp.cleanup()

    def test_retry_skips_received_sections_and_preserves_metadata(self):
        api = FakeApi()
        self.assertEqual(submit_batch(self.batch, api.post), 2)
        self.assertEqual(submit_batch(self.batch, api.post), 2)
        self.assertEqual(len(api.posts), 2)
        self.assertEqual(api.posts[0]["technicianName"], "Ryan")
        self.assertEqual(api.posts[0]["date"], "2026-09-17")
        self.assertEqual(api.posts[0]["visitType"], "Tracking")
        self.assertEqual(api.posts[0]["sourceKind"], "scanner")
        self.assertEqual(api.posts[0]["facilityNames"], [{"name": "Alpha", "popIn": False}])
        self.assertRegex(api.posts[0]["filename"], r"^Tracking Ryan 091726 Alpha \d{5}\.pdf$")
        self.assertEqual(api.posts[1]["visitType"], "Delivery")
        self.assertTrue(api.posts[1]["filename"].startswith("Delivery Ryan 091726 Beta "))

    def test_discard_removes_batch_and_only_an_intake_source(self):
        source = Path(self.temp.name) / "source.pdf"
        discard_batch(self.batch)
        self.assertFalse(self.batch.parent.exists())
        self.assertTrue(source.exists())
        batch = export_cst_batch(source, None, date(2026, 9, 17), [(1, 4, "Beta", "Tracking")], ["Beta"],
                                 [1, 2, 3, 4], {}, Path(self.temp.name) / "out", technician="Ryan") / "batch.json"
        with patch("app.cst_submission.INTAKE_DIR", source.parent):
            discard_batch(batch)
        self.assertFalse(source.exists())

    def test_lost_response_retries_with_the_same_id_and_payload(self):
        api = FakeApi()
        api.lose_response = True
        with self.assertRaises(TimeoutError):
            submit_batch(self.batch, api.post)
        self.assertEqual(submit_batch(self.batch, api.post), 2)
        self.assertEqual(len(api.receipts), 2)
        self.assertEqual(api.posts[0], api.posts[1])

    def test_payload_mismatch_never_marked_received(self):
        api = FakeApi()
        def post(payload):
            status, result = api.post(payload)
            return status, {**result, "payloadMismatch": True}
        with self.assertRaises(RuntimeError):
            submit_batch(self.batch, post)
        state = json.loads(self.batch.with_name("submission.json").read_text())
        self.assertFalse(state["jobs"][0]["received"])

    def test_http_transport_against_local_mock(self):
        api = FakeApi()
        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass
            def do_POST(self):
                body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
                code, result = api.post(body)
                data = json.dumps(result).encode()
                self.send_response(code)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            with patch("app.cst_submission.API_BASE", f"http://127.0.0.1:{server.server_port}"):
                self.assertEqual(submit_batch(self.batch), 2)
            self.assertEqual(len(api.posts), 2)
        finally:
            server.shutdown()
            server.server_close()
            thread.join()
