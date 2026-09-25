import json
import os
import tempfile
import threading
import unittest
from datetime import date, datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import fitz
from PySide6.QtWidgets import QApplication, QPushButton, QWidget, QVBoxLayout, QLineEdit
from PySide6.QtCore import QDate, Qt
from PySide6.QtTest import QTest

from app.cst_editor import CstEditorWindow, FacilityComboBox
from app.cst_workflow import export_cst_batch, fetch_facilities, parse_facilities
from app.outlook_intake import (IntakeStore, MAILBOX, MAPI_E_NOT_FOUND, extra_body_text, prune_left_inbox,
                                scan_outlook)


def make_pdf(path):
    with fitz.open() as doc:
        for number in range(1, 5):
            doc.new_page().insert_text((50, 50), f"Synthetic CST test page {number}")
        doc.save(path)


class CstWorkflowTests(unittest.TestCase):
    def test_facility_labels_and_bom(self):
        for encoding in ("utf-8-sig", "utf-16"):
            raw = "Alpha / Long Name\n\nBeta & Co.\nAlpha / Long Name\n".encode(encoding)
            self.assertEqual(parse_facilities(raw), ["Alpha / Long Name", "Beta & Co."])

    def test_facilities_fall_back_to_last_download_when_offline(self):
        with tempfile.TemporaryDirectory() as folder:
            cache = Path(folder) / "Facilities.txt"
            with patch("app.cst_workflow.urlopen", side_effect=OSError("offline")):
                with self.assertRaises(OSError):
                    fetch_facilities(cache)
                cache.write_bytes(b"Alpha\n")
                self.assertEqual(fetch_facilities(cache), ["Alpha"])

    def test_export_exact_names_pages_rotations_and_duplicate_facilities(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input.pdf"
            make_pdf(source)
            original = source.read_bytes()
            batch = export_cst_batch(source, None, date(2026, 9, 17),
                [(1, 2, "Alpha / Name", "Tracking"), (3, 4, "Alpha / Name", "Delivery")],
                ["Alpha / Name"], [4, 3, 2, 1], {4: 90}, root / "output", technician="Ryan")
            manifest = json.loads((batch / "batch.json").read_text())
            self.assertEqual(manifest["document_date"], "2026-09-17")
            self.assertEqual(manifest["technician"], "Ryan")
            self.assertEqual(manifest["sections"][0]["facility"], "Alpha / Name")
            self.assertEqual(manifest["sections"][0]["source_pages"], [4, 3])
            self.assertEqual([row["visit_type"] for row in manifest["sections"]], ["Tracking", "Delivery"])
            files = [batch / row["file"] for row in manifest["sections"]]
            self.assertNotEqual(files[0], files[1])
            with fitz.open(files[0]) as doc:
                self.assertEqual(len(doc), 2)
                self.assertIn("page 4", doc[0].get_text())
                self.assertEqual(doc[0].rotation, 90)
            self.assertEqual(source.read_bytes(), original)

    def test_failed_export_does_not_publish_partial_batch(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input.pdf"
            make_pdf(source)
            with patch("app.cst_workflow.pdf_ops.split_pdf", side_effect=OSError("disk full")):
                with self.assertRaises(OSError):
                    export_cst_batch(source, None, date.today(), [(1, 4, "Alpha", "Tracking")],
                                     ["Alpha"], [1, 2, 3, 4], {}, root / "out", technician="Ryan")
            self.assertEqual(list((root / "out").iterdir()), [])
            self.assertTrue(source.exists())

    def test_invalid_facility_rejected_before_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            with self.assertRaises(ValueError):
                export_cst_batch(root / "missing.pdf", None, date.today(),
                    [(1, 1, "Approximate", "Tracking")], ["Exact"], [1], {}, root / "out", technician="Ryan")
            self.assertFalse((root / "out").exists())

    def test_signature_only_and_extra_message(self):
        signature = "Sent from my T-Mobile 5G Device\vGet Outlook for Android <https://aka.ms/AAb9ysg>"
        self.assertEqual(extra_body_text(signature), "")
        self.assertEqual(extra_body_text("Please check these.\r\n" + signature), "Please check these.")
        self.assertEqual(extra_body_text("A different signature"), "A different signature")

    def test_queue_persists_and_deduplicates(self):
        with tempfile.TemporaryDirectory() as folder:
            store = IntakeStore(Path(folder))
            store.add("id", Path(folder) / "test.pdf", "note", "James", "entry", "store")
            store.add("id", Path(folder) / "test.pdf", "note", "James", "entry", "store")
            self.assertEqual(len(IntakeStore(Path(folder)).pending()), 1)
            store.complete("id")
            self.assertEqual(store.pending(), [])
            self.assertTrue(store.contains("id"))
            self.assertEqual(store.email("id"), ("entry", "store"))

    def test_fake_outlook_download_once_and_sender_filter(self):
        with tempfile.TemporaryDirectory() as folder:
            store = IntakeStore(Path(folder))
            class Attachment:
                FileName = "../../unsafe.pdf"
                saves = 0
                def SaveAsFile(self, path):
                    self.saves += 1
                    Path(path).write_bytes(b"synthetic attachment")
            attachment = Attachment()
            class Accessor:
                def GetProperty(self, name):
                    return "ryanmcbride30377@gmail.com" if "5D01001F" in name else "message-id"
            message = SimpleNamespace(Class=43, PropertyAccessor=Accessor(), EntryID="entry",
                ReceivedTime=datetime.now(timezone.utc), Body="Extra note",
                Attachments=SimpleNamespace(Count=1, Item=lambda index: attachment))
            class Items(list):
                def Restrict(self, query):
                    return self
                def Sort(self, key, descending):
                    pass
            inbox = SimpleNamespace(StoreID="office", Items=Items([message]))
            namespace = SimpleNamespace(Accounts=[SimpleNamespace(SmtpAddress=MAILBOX,
                DeliveryStore=SimpleNamespace(GetDefaultFolder=lambda _: inbox))])
            since = datetime(2026, 1, 1, tzinfo=timezone.utc)
            for _ in range(2):
                scan_outlook(namespace, store, since, threading.Event())
            self.assertEqual(attachment.saves, 1)
            self.assertEqual(len(store.pending()), 1)
            self.assertEqual(Path(store.pending()[0][1]).parent, Path(folder))
            self.assertEqual(store.pending()[0][2:], ("Extra note", "Ryan"))
            # A different sender is ignored. (The date cutoff is Outlook's Restrict.)
            message.PropertyAccessor = SimpleNamespace(GetProperty=lambda _: "other@example.com")
            scan_outlook(namespace, store, since, threading.Event())
            self.assertEqual(attachment.saves, 1)

    def test_dismiss_deletes_copy_and_never_requeues(self):
        with tempfile.TemporaryDirectory() as folder:
            store = IntakeStore(Path(folder))
            pdf = Path(folder) / "dup.pdf"
            pdf.write_bytes(b"x")
            store.add("dup", pdf, "", "James", "entry", "store")
            store.dismiss("dup")
            self.assertFalse(pdf.exists())
            self.assertEqual(store.pending(), [])
            self.assertTrue(store.contains("dup"))

    def test_emails_that_left_the_inbox_are_dropped(self):
        with tempfile.TemporaryDirectory() as folder:
            store = IntakeStore(Path(folder))
            for key in ("inbox", "deleted", "moved", "busy"):
                store.add(key, Path(folder) / f"{key}.pdf", "", "James", key, "store")
            inbox = SimpleNamespace(EntryID="INBOX")
            def get_item(entry_id, store_id):
                if entry_id == "moved":
                    raise Exception(-2147352567, "Exception occurred.",
                                    (4096, "Microsoft Outlook", "cannot be found", None, 0, MAPI_E_NOT_FOUND), None)
                if entry_id == "busy":
                    raise Exception(-2147352567, "Exception occurred.", (4096, "Microsoft Outlook", "busy", None, 0, -1), None)
                folder_id = "INBOX" if entry_id == "inbox" else "DELETED"
                return SimpleNamespace(Parent=SimpleNamespace(EntryID=folder_id))
            prune_left_inbox(SimpleNamespace(GetItemFromID=get_item), store, inbox)
            self.assertEqual([row[0] for row in store.pending()], ["inbox", "busy"])

    def test_dismiss_button_drops_active_intake(self):
        from app.tray_runtime import TrayRuntime
        runtime = object.__new__(TrayRuntime)
        runtime.active_intake_id = "dup"
        runtime.failed_intake_id = None
        calls = []
        runtime.intake_store = SimpleNamespace(complete=lambda k: calls.append(("complete", k)),
                                               dismiss=lambda k: calls.append(("dismiss", k)))
        runtime.cst_window = SimpleNamespace(set_aside=lambda: calls.append("set_aside"), metadata=None,
                                             submitting=False, hide=lambda: calls.append("hide"))
        runtime.watch_action = SimpleNamespace(isChecked=lambda: False)
        runtime._dismiss_intake()
        self.assertEqual(calls, [("complete", "dup"), "set_aside", "hide", ("dismiss", "dup")])
        self.assertIsNone(runtime.active_intake_id)

    def test_intake_does_not_replace_active_manual_work(self):
        from app.tray_runtime import TrayRuntime
        runtime = object.__new__(TrayRuntime)
        runtime.active_intake_id = None
        runtime.cst_window = SimpleNamespace(metadata=object(), submitting=False)
        # No store is installed: reaching the queue at all would fail this test.
        runtime._open_next_intake()
        runtime.cst_window = None
        runtime.active_intake_id = "already-open"
        runtime._open_next_intake()


class CstEditorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_tab_accepts_first_match_and_moves_focus(self):
        window = QWidget()
        layout = QVBoxLayout(window)
        combo = FacilityComboBox(["Alpha Care", "Alpha Village", "Beta"])
        following = QLineEdit()
        layout.addWidget(combo)
        layout.addWidget(following)
        window.show()
        window.activateWindow()
        self.app.processEvents()
        combo.setFocus()
        QTest.keyClicks(combo.lineEdit(), "alp")
        self.app.processEvents()
        QTest.keyClick(combo.lineEdit(), Qt.Key_Tab)
        self.app.processEvents()
        self.assertEqual(combo.currentText(), "Alpha Care")
        self.assertTrue(following.hasFocus())
        window.close()
        window.deleteLater()
        self.app.processEvents()

    def test_tab_keeps_highlighted_match_and_unmatched_input(self):
        combo = FacilityComboBox(["Alpha Care", "Alpha Village"])
        combo.setEditText("unknown")
        combo.accept_suggestion()
        self.assertEqual(combo.currentText(), "unknown")
        combo.setEditText("alp")
        combo.completer().setCompletionPrefix("alp")
        combo.show()
        combo.completer().complete()
        popup = combo.completer().popup()
        popup.setCurrentIndex(combo.completer().completionModel().index(1, 0))
        combo.accept_suggestion()
        self.assertEqual(combo.currentText(), "Alpha Village")
        combo.close()
        combo.deleteLater()
        self.app.processEvents()

    def test_facility_choices_follow_section_start_pages(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "synthetic.pdf"
            make_pdf(source)
            window = CstEditorWindow()
            window.bring_to_front = lambda: None
            with patch("app.cst_editor.fetch_facilities", return_value=["Alpha", "Beta"]):
                self.assertTrue(window.load_pdf(source))
            for _ in range(15):
                self.app.processEvents()
            window.show()
            window._toggle_split_start(3)
            self.app.processEvents()
            self.assertTrue(window.facility_inputs[3].hasFocus())
            window.document_date.setDate(QDate(2026, 9, 17))
            window.findChild(QPushButton, "date_step_-1").click()
            self.assertEqual(window.document_date.date(), QDate(2026, 9, 16))
            window.facility_inputs[1].setEditText("Alpha")
            window.facility_inputs[3].setEditText("Beta")
            window.page_order = [3, 4, 1, 2]
            window.split_starts = {1}
            window._refresh_sections_ui()
            self.assertEqual(window.facility_inputs[3].currentText(), "Beta")
            self.assertEqual(window.facility_inputs[1].currentText(), "Alpha")
            window._clear_loaded_pdf()
            self.assertEqual(window.section_facilities, {})
            window.close()
            window.deleteLater()
            self.app.processEvents()


if __name__ == "__main__":
    unittest.main()
