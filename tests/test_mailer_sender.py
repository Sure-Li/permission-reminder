from datetime import date, datetime
from email import policy
from email.parser import BytesParser
from io import BytesIO
from pathlib import Path
from queue import Queue
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from openpyxl import load_workbook

from permission_reminder.core.checker import ReminderItem
from permission_reminder.infra.config_store import AppConfig, SmtpConfig
from permission_reminder.mailer import outbox
from permission_reminder.mailer.sender import (
    EmailSender, TABLE_COLUMNS, build_excel_attachment, build_mime_message,
    render_summary_email, reset_auth_block,
)
from permission_reminder.ui import main_window


def sample_items():
    return [
        ReminderItem("00123", '张三 <script>"x"</script> &', "资管权限", "私募资管权限",
                     date(2026, 10, 21), date(2026, 10, 14), True),
        ReminderItem("00123", '张三 <script>"x"</script> &', "专业投资者", "专业投资者权限",
                     date(2026, 11, 1), date(2026, 10, 25), False),
        ReminderItem("=1+1", "=2+2", "私募基金", "私募基金权限",
                     date(2026, 10, 22), date(2026, 10, 15), False),
    ]


def parse_message(content):
    return BytesParser(policy=policy.default).parsebytes(content)


class MailerSummaryTests(unittest.TestCase):
    def setUp(self):
        self.items = sample_items()
        self.attachment = build_excel_attachment(self.items, datetime(2026, 10, 14, 1, 1, 2))

    def test_summary_groups_permissions_and_preserves_trigger_and_backfill(self):
        html, plain = render_summary_email(self.items, date(2026, 10, 16))
        self.assertEqual(html.count("<tr>"), 2)
        self.assertEqual(html.count("<th>"), 5)
        for header in TABLE_COLUMNS:
            self.assertIn(header, html)
            self.assertIn(header, plain)
        self.assertIn("2026-11-01", html)
        self.assertIn("2026-10-21", html)
        self.assertIn("【补发】", html)
        self.assertIn("原提醒时点 2026年10月14日", html)
        self.assertIn("原提醒时点：2026年10月14日", plain)
        self.assertNotIn("<script>", html)
        self.assertIn("&lt;script&gt;&quot;x&quot;&lt;/script&gt; &amp;", html)

    def test_excel_has_matching_five_columns_and_text_not_formulas(self):
        name, content = self.attachment
        self.assertEqual(name, "权限到期20261014010102.xlsx")
        workbook = load_workbook(BytesIO(content), data_only=False)
        try:
            sheet = workbook["权限到期提醒"]
            self.assertEqual(tuple(cell.value for cell in sheet[1]), TABLE_COLUMNS)
            self.assertEqual(sheet.max_row, 3)
            self.assertEqual(sheet["A2"].value, "00123")
            self.assertEqual(sheet["C2"].value, "2026-11-01")
            self.assertEqual(sheet["E2"].value, "2026-10-21")
            self.assertEqual(sheet["D3"].value, "2026-10-22")
            self.assertEqual(sheet["A3"].value, "=1+1")
            self.assertEqual(sheet["B3"].value, "=2+2")
            self.assertEqual(sheet["A3"].data_type, "s")
            self.assertEqual(sheet["B3"].data_type, "s")
        finally:
            workbook.close()

    def test_mime_contains_alternative_body_and_decodable_excel(self):
        html, plain = render_summary_email(self.items, date(2026, 10, 16))
        message = parse_message(build_mime_message(
            "sender@example.com", ["to@example.com"], "提醒邮件", html, plain,
            attachments=[self.attachment]).as_bytes())
        self.assertEqual(message.get_content_subtype(), "mixed")
        self.assertEqual(message.get_body(preferencelist=("plain",)).get_content(), plain)
        self.assertEqual(message.get_body(preferencelist=("html",)).get_content(), html)
        attachment, = message.iter_attachments()
        self.assertEqual(attachment.get_filename(), self.attachment[0])
        self.assertEqual(attachment.get_payload(decode=True), self.attachment[1])
        self.assertEqual(attachment.get_content_type(),
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    def test_exported_eml_contains_excel_and_old_export_interface_still_works(self):
        with TemporaryDirectory() as temp, patch.object(outbox, "DATA_DIR", Path(temp)):
            eml, html = outbox.export_mail(["to@example.com"], "提醒", "<p>表格</p>", "表格",
                                           attachments=[self.attachment])
            attachment, = parse_message(eml.read_bytes()).iter_attachments()
            self.assertEqual(attachment.get_payload(decode=True), self.attachment[1])
            self.assertIn("<p>表格</p>", html.read_text(encoding="utf-8"))
            eml_only = outbox.export_eml(["to@example.com"], "提醒", "h", "p",
                                         attachments=[self.attachment])
            self.assertEqual(len(list(parse_message(eml_only.read_bytes()).iter_attachments())), 1)
            old_eml, _ = outbox.export_mail(["to@example.com"], "旧接口", "h", "p")
            old_message = parse_message(old_eml.read_bytes())
            self.assertEqual(old_message.get_content_subtype(), "alternative")
            self.assertEqual(list(old_message.iter_attachments()), [])


class MailerDeliveryTests(unittest.TestCase):
    def setUp(self):
        reset_auth_block()
        self.items = sample_items()
        self.cfg = AppConfig(
            smtp=SmtpConfig(username="sender@example.com", password="test-only",
                            recipients=["to@example.com"]),
            excel_path="sample.xlsx", auto_sync_holidays=False)

    def test_smtp_optional_attachment_and_original_send_interface(self):
        sender = EmailSender(self.cfg.smtp)
        attachment = build_excel_attachment(self.items)
        with patch.object(sender, "_smtp_send") as send:
            sender.send(self.cfg.smtp.recipients, "提醒", "h", "p", attachments=[attachment])
            message = parse_message(send.call_args.args[1].as_bytes())
            self.assertEqual(next(message.iter_attachments()).get_payload(decode=True), attachment[1])
            sender.send(self.cfg.smtp.recipients, "原调用", "h", "p")
            self.assertEqual(send.call_args.args[1].get_content_subtype(), "alternative")

    def test_full_check_passes_attachment_to_smtp_export_and_client_routes(self):
        for mode, export_only in (("smtp", False), ("smtp", True), ("client", False)):
            with self.subTest(mode=mode, export_only=export_only):
                self.cfg.send_mode = mode
                fake = SimpleNamespace(_ui_queue=Queue(), _sent_store=Mock(),
                                       _export_and_handoff=Mock())
                fake._sent_store.filter_unsent.return_value = self.items
                with patch.object(main_window, "read_clients", return_value=[]), \
                        patch.object(main_window, "find_due_items", return_value=self.items), \
                        patch.object(main_window, "create_calendar", return_value=Mock()), \
                        patch.object(main_window, "export_mail", return_value=(Path("mail.eml"), Path("mail.html"))) as export, \
                        patch.object(EmailSender, "_smtp_send") as send:
                    main_window.MainWindow._do_check(fake, self.cfg, export_only=export_only)
                    if export_only:
                        self.assertEqual(len(export.call_args.kwargs["attachments"]), 1)
                        fake._sent_store.mark_sent.assert_not_called()
                    elif mode == "client":
                        self.assertEqual(len(fake._export_and_handoff.call_args.kwargs["attachments"]), 1)
                        fake._sent_store.mark_sent.assert_not_called()
                    else:
                        self.assertEqual(len(list(parse_message(send.call_args.args[1].as_bytes()).iter_attachments())), 1)
                        self.assertEqual(fake._sent_store.mark_sent.call_count, len(self.items))

    def test_client_handoff_forces_eml_even_when_foxmail_prefers_mailto(self):
        self.cfg.foxmail_mode = "mailto"
        fake = SimpleNamespace(_ui_queue=Queue())
        html, plain = render_summary_email(self.items, date(2026, 10, 16))
        with TemporaryDirectory() as temp, patch.object(outbox, "DATA_DIR", Path(temp)), \
                patch.object(main_window, "resolve_foxmail", return_value=("Foxmail.exe", "")), \
                patch.object(main_window, "open_foxmail_compose", return_value=[]) as foxmail, \
                patch.object(main_window, "open_in_mail_client") as default_client:
            main_window.MainWindow._export_and_handoff(fake, self.cfg, self.items, "提醒", html, plain)
            self.assertEqual(foxmail.call_args.args[1], "eml")
            path = Path(foxmail.call_args.args[2])
            self.assertEqual(len(list(parse_message(path.read_bytes()).iter_attachments())), 1)
            default_client.assert_not_called()
            self.assertEqual(fake._ui_queue.get_nowait()["type"], "confirm_export")

    def test_default_client_handoff_receives_eml_path_with_attachment(self):
        fake = SimpleNamespace(_ui_queue=Queue())
        html, plain = render_summary_email(self.items, date(2026, 10, 16))
        with TemporaryDirectory() as temp, patch.object(outbox, "DATA_DIR", Path(temp)), \
                patch.object(main_window, "resolve_foxmail", return_value=("", "")), \
                patch.object(main_window, "describe_mail_client", return_value="测试客户端"), \
                patch.object(main_window, "open_in_mail_client", return_value=True) as open_client:
            main_window.MainWindow._export_and_handoff(fake, self.cfg, self.items, "提醒", html, plain)
            path = open_client.call_args.kwargs["eml_path"]
            self.assertEqual(len(list(parse_message(path.read_bytes()).iter_attachments())), 1)

    def test_opening_attachment_mail_uses_eml_and_no_mailto_fallback(self):
        with TemporaryDirectory() as temp:
            path = Path(temp) / "mail.eml"
            path.write_text("test", encoding="utf-8")
            with patch.object(outbox.sys, "platform", "win32"), \
                    patch.object(outbox, "_query_assoc", return_value="Mail.exe"), \
                    patch.object(outbox.os, "startfile", create=True) as start, \
                    patch.object(outbox, "build_mailto") as mailto:
                self.assertTrue(outbox.open_in_mail_client(["to@example.com"], "提醒", "正文", eml_path=path))
                start.assert_called_once_with(str(path.resolve()))
                mailto.assert_not_called()
            with patch.object(outbox.sys, "platform", "win32"), \
                    patch.object(outbox, "_query_assoc", return_value=""), \
                    patch.object(outbox.os, "startfile", create=True) as start:
                self.assertFalse(outbox.open_in_mail_client(["to@example.com"], "提醒", "正文", eml_path=path))
                start.assert_not_called()


if __name__ == "__main__":
    unittest.main()
