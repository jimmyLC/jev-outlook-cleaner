"""Offline tests for the local browser boundary; no mail or Jev calls."""

import json
from io import BytesIO
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import web_app


class WebTests(unittest.TestCase):
    def test_review_path_rejects_traversal_and_symlink(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(web_app, "REVIEWS", Path(folder)):
            with self.assertRaises(ValueError):
                web_app.review_path("../.env")
            target = Path(folder) / "review-20261021-090000-1.json"
            target.symlink_to(Path(folder) / ".env")
            with self.assertRaises(ValueError):
                web_app.review_path(target.name)

    def test_dashboard_only_displays_counts_and_review_name(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(web_app, "REVIEWS", Path(folder)):
            path = Path(folder) / "review-20261021-090000-1.json"
            path.write_text(json.dumps({"items": [{"decision": "keep", "subject": "private subject",
                                                    "body": "private body"}]}))
            page = web_app.dashboard()
            self.assertIn("保留 1", page)
            self.assertNotIn("private subject", page)
            self.assertNotIn("private body", page)

    def test_post_requires_browser_origin_and_csrf(self):
        handler = web_app.Handler.__new__(web_app.Handler)
        handler.server = SimpleNamespace(server_port=8765)
        handler.path = "/api/scan"
        body = json.dumps({"days": 30, "limit": 2}).encode()
        handler.rfile = BytesIO(body)
        handler.headers = {"Host": "127.0.0.1:8765", "Content-Type": "application/json",
                           "Content-Length": str(len(body)), "Origin": "https://example.com",
                           "X-CSRF-Token": web_app.CSRF}
        with patch.object(handler, "send_bytes") as reply, patch.object(web_app.JOBS, "start", return_value=True) as start:
            handler.do_POST()
            self.assertEqual(reply.call_args.args[0], 403)
            start.assert_not_called()
            handler.headers["Origin"] = "http://127.0.0.1:8765"
            handler.headers["X-CSRF-Token"] = "invalid"
            handler.do_POST()
            self.assertEqual(reply.call_args.args[0], 403)
            start.assert_not_called()
            handler.headers["X-CSRF-Token"] = web_app.CSRF
            handler.do_POST()
            self.assertEqual(reply.call_args.args[0], 202)
            start.assert_called_once_with("scan", ["scan", "--older-than-days", "30", "--limit", "2"])


if __name__ == "__main__":
    unittest.main()
