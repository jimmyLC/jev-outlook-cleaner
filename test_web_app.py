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

    def test_login_job_exposes_only_validated_device_code(self):
        class FakeProc:
            returncode = 0
            stdout = iter([
                'JEV_DEVICE_LOGIN {"url": "https://microsoft.com/devicelogin", "code": "ABCD1234"}\n',
                "若要登入，請開啟頁面並輸入代碼 ABCD1234\n",
                "Outlook 登入完成。\n",
            ])
            def wait(self):
                return 0
            def kill(self):
                pass
        jobs = web_app.Jobs()
        seen = []
        with patch.object(web_app.subprocess, "Popen", return_value=FakeProc()) as popen:
            self.assertTrue(jobs.start_login())
            for _ in range(100):
                if not jobs.active:
                    break
                import time; time.sleep(0.01)
        command = popen.call_args.args[0]
        self.assertEqual(command[-1], "login")
        self.assertNotIn("JEV_NONINTERACTIVE", popen.call_args.kwargs["env"])
        self.assertEqual(jobs.snapshot(), {"state": "done", "operation": "login"})

    def test_device_login_line_rejects_foreign_hosts_and_bad_codes(self):
        line = lambda url, code: "JEV_DEVICE_LOGIN " + json.dumps({"url": url, "code": code})
        self.assertIsNone(web_app.parse_device_login(line("https://evil.example/link", "ABCD1234")))
        self.assertIsNone(web_app.parse_device_login(line("https://microsoft.com/devicelogin", "<b>x</b>")))
        self.assertIsNone(web_app.parse_device_login("JEV_DEVICE_LOGIN not-json"))
        self.assertEqual(web_app.parse_device_login(line("https://microsoft.com/devicelogin", "ABCD1234")),
                         ("https://microsoft.com/devicelogin", "ABCD1234"))

    def test_login_is_refused_while_another_job_runs(self):
        jobs = web_app.Jobs()
        jobs.active = True
        self.assertFalse(jobs.start_login())


if __name__ == "__main__":
    unittest.main()
