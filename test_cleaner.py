"""Offline checks: no Outlook or Jev network calls are made."""
import copy
import json
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import jev_outlook_cleaner as cleaner


def message():
    return {"id": "message-1", "subject": "舊優惠活動", "sender": {"emailAddress": {
        "name": "Shop", "address": "shop@example.com"}}, "receivedDateTime": "2026-01-01T00:00:00Z",
        "internetMessageId": "<one@example.com>", "body": {"content": "Expired sale", "contentType": "text"},
        "importance": "normal", "isRead": True, "flag": {"flagStatus": "notFlagged"}, "hasAttachments": False}


def review(msg):
    return {"schema_version": 2, "account_id": "account-1", "account_address": "me@example.com", "items": [{
        "id": msg["id"], "subject": msg["subject"], "sender_name": "Shop", "sender_address": "shop@example.com",
        "received_at": msg["receivedDateTime"], "internet_message_id": msg["internetMessageId"],
        "snapshot": cleaner.snapshot(msg), "decision": "delete_candidate", "confidence": 0.99,
        "excerpt": "<script>alert(1)</script>"}]}


class WorkflowTests(unittest.TestCase):
    def test_web_worker_never_starts_hidden_device_sign_in(self):
        with tempfile.TemporaryDirectory() as folder:
            app = Mock()
            app.get_accounts.return_value = []
            with patch.object(cleaner, "ROOT", Path(folder)), \
                    patch.object(cleaner, "CACHE_PATH", Path(folder) / "missing-cache"), \
                    patch.object(cleaner, "graph_client_id", return_value="example-client"), \
                    patch.object(cleaner.msal, "PublicClientApplication", return_value=app), \
                    patch.dict(os.environ, {"JEV_NONINTERACTIVE": "1"}):
                with self.assertRaisesRegex(RuntimeError, "login"):
                    cleaner.access_token()
            app.initiate_device_flow.assert_not_called()

    def test_protected_messages_never_call_jev(self):
        for change in ({"hasAttachments": True}, {"importance": "high"}, {"flag": {"flagStatus": "flagged"}}):
            msg = message() | change
            client = Mock()
            self.assertEqual(cleaner.classify(client, msg)["decision"], "keep")
            client.system_one.assert_not_called()

    def test_low_confidence_delete_goes_to_review(self):
        client = Mock()
        client.system_one.return_value = SimpleNamespace(choices={
            "triage": SimpleNamespace(choice="delete_candidate", confidence=0.4,
                                       probabilities={"delete_candidate": 0.6, "keep": 0.2, "review": 0.2}),
            "message_type": SimpleNamespace(choice="bulk_or_automated", confidence=0.99),
        })
        self.assertEqual(cleaner.classify(client, message())["decision"], "review")

    def test_content_change_invalidates_snapshot(self):
        first = message()
        second = copy.deepcopy(first)
        second["body"]["content"] = "Please pay this invoice"
        self.assertNotEqual(cleaner.snapshot(first), cleaner.snapshot(second))

    def test_preview_escapes_email_html(self):
        data = review(message())
        data["items"][0]["subject"] = "<script>alert(1)</script>"
        output = cleaner.render_review(data)
        self.assertNotIn("<script>alert(1)</script>", output)
        self.assertIn("&lt;script&gt;alert(1)&lt;/script&gt;", output)

    def test_preview_has_keep_checkbox_for_every_decision(self):
        data = review(message())
        second = copy.deepcopy(data["items"][0])
        second.update(id="message-2", decision="keep", subject="Suggested keep")
        data["items"].append(second)
        output = cleaner.render_review(data)
        self.assertEqual(output.count('class="keep-toggle"'), 2)
        self.assertIn("保留清單」分類", output)
        self.assertIn('data-review-key="' + cleaner.review_selection_key(data) + '"', output)
        self.assertNotIn('value="message-1"', output)

    def test_already_moved_items_are_not_selectable_again(self):
        data = review(message())
        data["items"][0]["moved_to_deleted_items"] = True
        output = cleaner.render_review(data)
        self.assertNotIn('class="keep-toggle"', output)
        self.assertIn("已移至 Outlook「刪除的郵件」", output)

    def test_selection_key_is_bound_to_review_identity(self):
        data = review(message())
        original = cleaner.review_selection_key(data)
        changed = copy.deepcopy(data)
        changed["items"][0]["id"] = "another-id"
        self.assertNotEqual(original, cleaner.review_selection_key(changed))

    def test_generic_no_reply_is_a_signal_not_an_automatic_rule(self):
        self.assertIsNone(cleaner.learned_delete_rule(
            "Monthly newsletter", "no-reply@news.example", "2015-01-01T00:00:00Z"
        ))

    def test_auto_delete_requires_confident_bulk_and_triage(self):
        item = review(message())["items"][0]
        item.update(message_type="bulk_or_automated", message_type_confidence=0.9)
        self.assertTrue(cleaner.auto_deletion_candidate(item))
        item["message_type_confidence"] = 0.89
        self.assertFalse(cleaner.auto_deletion_candidate(item))
        item["message_type_confidence"] = 0.99
        item["confidence"] = 0.84
        self.assertFalse(cleaner.auto_deletion_candidate(item))

    def test_auto_apply_rejects_legacy_generic_no_reply_rule(self):
        item = review(message())["items"][0]
        item["rule"] = "超過 30 天的 no-reply 自動寄件信"
        self.assertFalse(cleaner.auto_deletion_candidate(item))

    def test_auto_apply_moves_high_confidence_candidate_without_prompt(self):
        msg = message()
        data = review(msg)
        data["items"][0].update(message_type="bulk_or_automated", message_type_confidence=0.99)
        graph = Mock()
        graph.profile.return_value = {"id": "account-1"}
        graph.current_inbox_message.return_value = msg
        graph.move_to_trash.return_value = {"id": "moved-1"}
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "review.json"
            path.write_text(json.dumps(data), encoding="utf-8")
            with patch.object(cleaner, "access_token", return_value="fake"), \
                    patch.object(cleaner, "Outlook", return_value=graph):
                result = cleaner.auto_apply(SimpleNamespace(review_file=str(path)))
        self.assertEqual(result, 0)
        graph.move_to_trash.assert_called_once_with("message-1")

    def run_apply(self, inputs, graph=None, data=None):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "review.json"
            path.write_text(json.dumps(data or review(message())), encoding="utf-8")
            with patch("builtins.input", side_effect=inputs), patch.object(cleaner, "access_token", return_value="fake"), \
                    patch.object(cleaner, "Outlook", return_value=graph) as factory:
                result = cleaner.apply(SimpleNamespace(review_file=str(path)))
                return result, factory

    def test_cancel_never_connects_to_outlook(self):
        result, factory = self.run_apply([""])
        self.assertEqual(result, 0)
        factory.assert_not_called()

    def test_changed_message_is_not_moved(self):
        graph = Mock()
        graph.profile.return_value = {"id": "account-1"}
        graph.current_inbox_message.return_value = message() | {"importance": "high"}
        result, _ = self.run_apply(["1", "TRASH 1"], graph)
        self.assertEqual(result, 1)
        graph.move_to_trash.assert_not_called()

    def test_wrong_account_is_not_moved(self):
        graph = Mock()
        graph.profile.return_value = {"id": "other-account"}
        with self.assertRaises(RuntimeError):
            self.run_apply(["1", "TRASH 1"], graph)
        graph.move_to_trash.assert_not_called()

    def test_confirmed_unchanged_message_moves_once(self):
        graph = Mock()
        graph.profile.return_value = {"id": "account-1"}
        graph.current_inbox_message.return_value = message()
        graph.move_to_trash.return_value = {"id": "moved-1"}
        result, _ = self.run_apply(["1", "TRASH 1"], graph)
        self.assertEqual(result, 0)
        graph.move_to_trash.assert_called_once_with("message-1")

    def run_apply_selection(self, selection, data=None, graph=None, keep_subjects=()):
        with tempfile.TemporaryDirectory() as folder:
            project = Path(folder)
            review_path = project / "review.json"
            data = data or review(message())
            review_path.write_text(json.dumps(data), encoding="utf-8")
            selection_path = project / "selection.json"
            selection_path.write_text(json.dumps(selection), encoding="utf-8")
            if keep_subjects:
                (project / "keep-list.json").write_text(
                    json.dumps({"version": 1, "exact_subjects": list(keep_subjects)}), encoding="utf-8"
                )
            with patch.object(cleaner, "PROJECT", project), \
                    patch.object(cleaner, "access_token", return_value="fake"), \
                    patch.object(cleaner, "Outlook", return_value=graph):
                result = cleaner.apply_selection(SimpleNamespace(
                    review_file=str(review_path), selection_file=str(selection_path)
                ))
                keep_data = json.loads((project / "keep-list.json").read_text(encoding="utf-8")) \
                    if (project / "keep-list.json").exists() else {}
                return result, keep_data

    def test_checkbox_selection_keeps_checked_and_deletes_unchecked_even_if_jev_says_keep(self):
        msg1 = message()
        msg2 = message() | {"id": "message-2", "subject": "Jev says keep",
                            "internetMessageId": "<two@example.com>"}
        data = review(msg1)
        data["items"].append({
            "id": msg2["id"], "subject": msg2["subject"], "sender_name": "Shop",
            "sender_address": "shop@example.com", "received_at": msg2["receivedDateTime"],
            "internet_message_id": msg2["internetMessageId"], "snapshot": cleaner.snapshot(msg2),
            "decision": "keep", "confidence": None,
        })
        selection = {"schema_version": 1, "review_key": cleaner.review_selection_key(data),
                     "confirmed": True, "keep_indexes": [1]}
        graph = Mock()
        graph.profile.return_value = {"id": "account-1"}
        graph.current_inbox_message.return_value = msg1
        graph.move_to_trash.return_value = {"id": "moved-1"}
        result, keep_data = self.run_apply_selection(selection, data, graph)
        self.assertEqual(result, 0)
        graph.move_to_trash.assert_called_once_with("message-1")
        self.assertIn("Jev says keep", keep_data["exact_subjects"])

    def test_apply_selection_rejects_mismatched_review_without_graph_calls(self):
        data = review(message())
        selection = {"schema_version": 1, "review_key": "wrong", "confirmed": True, "keep_indexes": []}
        graph = Mock()
        with self.assertRaises(RuntimeError):
            self.run_apply_selection(selection, data, graph)
        graph.profile.assert_not_called()

    def test_selection_file_cannot_reselect_an_already_moved_item(self):
        data = review(message())
        data["items"][0]["moved_to_deleted_items"] = True
        selection = {"schema_version": 1, "review_key": cleaner.review_selection_key(data),
                     "confirmed": True, "keep_indexes": [0]}
        graph = Mock()
        with self.assertRaises(RuntimeError):
            self.run_apply_selection(selection, data, graph)
        graph.profile.assert_not_called()

    def test_existing_keep_rules_override_unchecked_selection(self):
        msg = message()
        data = review(msg)
        selection = {"schema_version": 1, "review_key": cleaner.review_selection_key(data),
                     "confirmed": True, "keep_indexes": []}
        graph = Mock()
        graph.profile.return_value = {"id": "account-1"}
        result, keep_data = self.run_apply_selection(
            selection, data, graph, keep_subjects=[msg["subject"]]
        )
        self.assertEqual(result, 0)
        graph.current_inbox_message.assert_not_called()
        graph.move_to_trash.assert_not_called()
        self.assertEqual(keep_data["exact_subjects"], [msg["subject"]])


class SignalAndLogicTests(unittest.TestCase):
    def jev(self, triage="delete_candidate", tri_conf=0.99, kind="bulk_or_automated", kind_conf=0.99):
        client = Mock()
        client.system_one.return_value = SimpleNamespace(choices={
            "triage": SimpleNamespace(choice=triage, confidence=tri_conf, probabilities={}),
            "message_type": SimpleNamespace(choice=kind, confidence=kind_conf),
        })
        return client

    def test_security_and_receipt_mail_never_reach_jev(self):
        for subject in ("Your verification code", "訂單確認 #123", "Re: 週末聚餐", "您的電子發票"):
            client = Mock()
            result = cleaner.classify(client, message() | {"subject": subject})
            self.assertEqual(result["decision"], "keep", subject)
            client.system_one.assert_not_called()

    def test_jev_bulk_without_local_evidence_is_downgraded_to_review(self):
        result = cleaner.classify(self.jev(), message())
        self.assertEqual(result["decision"], "review")
        self.assertIn("缺少", result["note"])

    def test_jev_bulk_with_unsubscribe_evidence_stays_delete_candidate(self):
        msg = message()
        msg["body"]["content"] = "Big sale! Click here to unsubscribe"
        result = cleaner.classify(self.jev(), msg)
        self.assertEqual(result["decision"], "delete_candidate")
        self.assertIn("內文含取消訂閱", result["signals"]["bulk"])

    def test_list_headers_count_as_bulk_evidence(self):
        msg = message() | {"internetMessageHeaders": [{"name": "List-Unsubscribe", "value": "<mailto:x@y>"}]}
        self.assertEqual(cleaner.classify(self.jev(), msg)["decision"], "delete_candidate")

    def test_list_headers_do_not_change_the_move_time_snapshot(self):
        msg = message()
        with_headers = msg | {"internetMessageHeaders": [{"name": "X", "value": "1"}]}
        self.assertEqual(cleaner.snapshot(msg), cleaner.snapshot(with_headers))

    def test_auto_apply_refuses_item_with_empty_bulk_signals(self):
        item = review(message())["items"][0]
        item.update(message_type="bulk_or_automated", message_type_confidence=0.99,
                    signals={"bulk": [], "protect": []})
        self.assertFalse(cleaner.auto_deletion_candidate(item))
        item["signals"]["bulk"] = ["內文含取消訂閱"]
        self.assertTrue(cleaner.auto_deletion_candidate(item))

    def rules_in(self, rules):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        (Path(folder.name) / "delete-rules.json").write_text(json.dumps(rules), encoding="utf-8")
        patcher = patch.object(cleaner, "PROJECT", Path(folder.name))
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_generic_no_reply_rule_does_not_shadow_exact_sender_rule(self):
        self.rules_in({"sender_rules": [
            {"type": "no_reply_local_part"},
            {"email": "noreply@shop.example", "min_age_days": 30, "reason": "exact"},
        ]})
        self.assertEqual(cleaner.learned_delete_rule(
            "Promo", "noreply@shop.example", "2020-01-01T00:00:00Z"), "exact")

    def test_invalid_saved_regex_is_skipped_not_fatal(self):
        self.rules_in({"subject_patterns": [{"regex": "("}, {"regex": "Weekly .*", "reason": "ok"}]})
        self.assertEqual(cleaner.learned_delete_rule("Weekly digest"), "ok")

    def test_preview_shows_reason_and_formatted_date(self):
        data = review(message())
        data["items"][0].update(signals={"bulk": ["內文含取消訂閱"]}, confidence=0.93)
        output = cleaner.render_review(data)
        self.assertIn("群發佐證：內文含取消訂閱", output)
        self.assertIn("Jev 信心 93%", output)
        self.assertIn("2026-01-01 00:00", output)


if __name__ == "__main__":
    unittest.main()
