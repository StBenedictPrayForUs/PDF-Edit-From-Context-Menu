from __future__ import annotations

import argparse
import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QSettings, QTimer, Signal
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import QApplication, QFileDialog, QMenu, QMessageBox, QStyle, QSystemTrayIcon

from .editor_window import PdfEditorWindow
from .ipc import IpcServer
from .logging_utils import configure_logging
from .combine_flow import run_combine_dialog, run_convert_image_dialog
from .cst_editor import CstEditorWindow
from .outlook_intake import (INTAKE_DIR, SUBJECT_FACILITY_TECHNICIANS, TECHNICIANS, IntakeStore, file_email,
                             poll_outlook)


class MessageBridge(QObject):
    received = Signal(dict)
    intake_finished = Signal(str)


class TrayRuntime:
    def __init__(self, app: QApplication) -> None:
        self.app = app
        self.window: PdfEditorWindow | None = None
        self._busy_action: str | None = None
        self.cst_window: CstEditorWindow | None = None
        self.active_intake_id: str | None = None
        self.failed_intake_id: str | None = None
        self.settings = QSettings("RMRR", "PDF Splitter")
        self.intake_store = IntakeStore(INTAKE_DIR)
        self.intake_worker: threading.Thread | None = None
        self.intake_stop = threading.Event()

        self.bridge = MessageBridge()
        self.bridge.received.connect(self._handle_ipc_message)
        self.bridge.intake_finished.connect(self._intake_finished)

        self.server = IpcServer(self.bridge.received.emit)
        if not self.server.start():
            raise RuntimeError("Another PDF Page Editor tray instance is already running.")
        logging.info("Tray runtime started and IPC server listening.")

        self.tray = QSystemTrayIcon(self._icon())
        self.tray.setToolTip("PDF Page Editor")
        self.tray.activated.connect(self._on_tray_activated)
        self.tray.setContextMenu(self._build_menu())
        self.tray.show()
        self.intake_timer = QTimer()
        self.intake_timer.setInterval(60_000)
        self.intake_timer.timeout.connect(self._poll_intake)
        self.intake_timer.start()
        if self.watch_action.isChecked():
            QTimer.singleShot(0, self._poll_intake)

    def _icon(self) -> QIcon:
        return self.app.style().standardIcon(QStyle.SP_FileDialogDetailedView)

    def _build_menu(self) -> QMenu:
        menu = QMenu()

        open_file_action = QAction("Open PDF...", menu)
        open_file_action.triggered.connect(self._pick_pdf)

        show_action = QAction("Show Window", menu)
        show_action.triggered.connect(self._show_window)

        quit_action = QAction("Quit", menu)
        quit_action.triggered.connect(self.shutdown)

        menu.addAction(open_file_action)
        menu.addAction(show_action)
        menu.addSeparator()
        cst_action = QAction("Open CST PDF (date + facilities)…", menu)
        cst_action.triggered.connect(self._pick_cst_pdf)
        menu.addAction(cst_action)
        show_cst = QAction("Show CST workspace / queued PDF", menu)
        show_cst.triggered.connect(self._show_cst)
        menu.addAction(show_cst)
        skip_cst = QAction("Set aside current queued PDF…", menu)
        skip_cst.triggered.connect(self._skip_intake)
        menu.addAction(skip_cst)
        self.watch_action = QAction("Automatically open emailed CST PDFs", menu)
        self.watch_action.setCheckable(True)
        self.watch_action.setChecked(self.settings.value("cst/watch_enabled", False, bool))
        self.watch_action.toggled.connect(self._toggle_intake)
        menu.addAction(self.watch_action)
        self.intake_status = QAction("CST intake is off", menu)
        self.intake_status.setEnabled(False)
        menu.addAction(self.intake_status)
        menu.addSeparator()
        menu.addAction(quit_action)
        return menu

    def _ensure_cst_window(self) -> CstEditorWindow:
        if self.cst_window is None:
            self.cst_window = CstEditorWindow()
            self.cst_window.batch_exported.connect(self._cst_exported)
            self.cst_window.submission_finished.connect(self._cst_submission_finished)
            self.cst_window.intake_dismissed.connect(self._dismiss_intake)
            self.cst_window.on_received = lambda source: (
                file_email(self.intake_store, source.stem) if source.parent == INTAKE_DIR else "")
        return self.cst_window

    def _cst_submission_finished(self, success: bool, message: str) -> None:
        if success:
            self.tray.showMessage("CST submitted", message)
        # A submission holds the window; the next queued PDF can open now.
        self._open_next_if_watching()

    def _pick_cst_pdf(self) -> None:
        self._ensure_cst_window().pick_pdf()

    def _show_cst(self) -> None:
        self.failed_intake_id = None
        self._open_next_intake()
        self._ensure_cst_window().bring_to_front()

    def _skip_intake(self) -> None:
        if self.cst_window and self.cst_window.submitting:
            self.cst_window.bring_to_front()
            return
        pending = self.intake_store.pending()
        if not pending:
            return
        key = self.active_intake_id or pending[0][0]
        answer = QMessageBox.question(None, "Set aside queued PDF?",
            "Leave this PDF saved locally and remove it from the automatic queue? "
            "Any unexported split choices will be discarded. The email stays in Outlook.",
            QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel)
        if answer != QMessageBox.Yes:
            return
        self.intake_store.complete(key)
        self._close_active_intake()

    def _dismiss_intake(self) -> None:
        if self.active_intake_id:
            key = self.active_intake_id
            self.intake_store.complete(key)
            self._close_active_intake()
            self.intake_store.dismiss(key)  # Deletes the copy now that the window released it.

    def _close_active_intake(self) -> None:
        if self.active_intake_id and self.cst_window:
            self.cst_window.set_aside()
        self.active_intake_id = None
        self.failed_intake_id = None
        self._open_next_if_watching()
        if self.cst_window and self.cst_window.metadata is None and not self.cst_window.submitting:
            self.cst_window.hide()  # Nothing left to work on.

    def _toggle_intake(self, enabled: bool) -> None:
        if enabled and not self.settings.value("cst/watch_since", "", str):
            answer = QMessageBox.question(None, "Enable automatic CST intake?",
                "While Outlook Classic is open, check the office mailbox every minute for new "
                f"PDFs from {' and '.join(TECHNICIANS)} and open them here?\n\n"
                "Today's emails are included; older ones will not open. Emails stay in the Inbox. "
                "If a PDF is already being edited, new arrivals wait in a local queue.",
                QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel)
            if answer != QMessageBox.Yes:
                self.watch_action.blockSignals(True)
                self.watch_action.setChecked(False)
                self.watch_action.blockSignals(False)
                return
            midnight = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
            self.settings.setValue("cst/watch_since", midnight.isoformat())
        self.settings.setValue("cst/watch_enabled", enabled)
        if enabled:
            self._poll_intake()
        else:
            self.intake_stop.set()
            self.intake_status.setText("CST intake paused • queued PDFs retained")

    def _poll_intake(self) -> None:
        if not self.watch_action.isChecked():
            return
        if self.intake_worker is not None and self.intake_worker.is_alive():
            return
        enabled_since = self.settings.value("cst/watch_since", "", str)
        if not enabled_since:  # The timer can fire while the enable prompt is still open.
            return
        self.intake_stop = threading.Event()
        stop = self.intake_stop
        self.intake_status.setText("Checking Outlook Classic…")
        def work():
            self.bridge.intake_finished.emit(poll_outlook(self.intake_store, enabled_since, stop))
        self.intake_worker = threading.Thread(target=work, name="CST Outlook intake", daemon=True)
        self.intake_worker.start()

    def _intake_finished(self, status: str) -> None:
        if not self.watch_action.isChecked() or self.intake_stop.is_set():
            return
        pending = self.intake_store.pending()
        self.intake_status.setText(f"{status} • {len(pending)} pending")
        if self.active_intake_id and self.active_intake_id not in {row[0] for row in pending}:
            # Its email left the Inbox, so the poll dropped it.
            self.tray.showMessage("CST PDF closed", "Its email is no longer in the Inbox, so it won't be processed.")
            self._dismiss_intake()  # Retries the delete the worker couldn't do while the PDF was open.
        else:
            self._open_next_intake()

    def _open_next_if_watching(self) -> None:
        if self.watch_action.isChecked():
            self._open_next_intake()

    def _open_next_intake(self) -> None:
        if self.active_intake_id:
            return
        if self.cst_window is not None and (self.cst_window.submitting or self.cst_window.metadata is not None):
            return
        pending = self.intake_store.pending()
        if not pending or pending[0][0] == self.failed_intake_id:
            return
        key, path, note, technician, subject = pending[0]
        window = self._ensure_cst_window()
        if not window.load_pdf(path):
            self.failed_intake_id = key
            self.intake_status.setText("Needs attention: queued PDF could not be opened. Set it aside to continue.")
            return
        self.active_intake_id = key
        window.intake_locked = True
        window.technician.setCurrentText(technician)
        # Ryan sends one facility per email; James's subjects list several.
        filled = window.prefill_from_subject(subject, facility=technician in SUBJECT_FACILITY_TECHNICIANS)
        window.intake_note.setText(f"{technician}’s email: " + (note or "No message beyond the signature.") + filled)
        if note:
            QMessageBox.information(window, f"{technician} included a message", note)

    def _cst_exported(self, batch: str) -> None:
        if self.active_intake_id:
            self.intake_store.complete(self.active_intake_id)
            self.active_intake_id = None
        if self.cst_window:
            self.cst_window.intake_locked = False
        QTimer.singleShot(0, self._open_next_if_watching)

    def shutdown(self) -> None:
        logging.info("Tray runtime shutting down.")
        self.intake_stop.set()
        self.server.stop()
        self.tray.hide()
        self.app.quit()

    def _ensure_window(self) -> PdfEditorWindow:
        if self.window is None:
            self.window = PdfEditorWindow()
        return self.window

    def _show_window(self) -> None:
        window = self._ensure_window()
        window.bring_to_front()

    def _pick_pdf(self) -> None:
        chosen, _ = QFileDialog.getOpenFileName(None, "Select PDF", "", "PDF Files (*.pdf)")
        if chosen:
            self.open_pdf(chosen)

    def open_pdf(self, path: str) -> None:
        resolved = str(Path(path).resolve())
        logging.info("Opening PDF: %s", resolved)
        window = self._ensure_window()
        try:
            loaded = window.load_pdf(resolved)
        except Exception:
            logging.exception("Unhandled error while opening PDF: %s", resolved)
            QMessageBox.critical(None, "PDF Page Editor Error", "Unexpected error while opening PDF.")
            return
        if not loaded:
            logging.warning("PDF load was canceled or failed: %s", resolved)

    def _handle_ipc_message(self, payload: dict) -> None:
        action = payload.get("action")
        if action == "open_pdf":
            raw_path = payload.get("path")
            if not raw_path:
                return

            path = Path(raw_path)
            if not path.exists():
                QMessageBox.warning(None, "PDF Not Found", f"Could not find:\n{raw_path}")
                return

            self.open_pdf(str(path))
            return

        if action == "combine_documents":
            self._run_single_action("combine", payload.get("paths", []), run_combine_dialog)
            return

        if action == "convert_images":
            self._run_single_action("convert-image", payload.get("paths", []), run_convert_image_dialog)
            return

    def _on_tray_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self._show_window()

    def _run_single_action(
        self,
        action_name: str,
        paths: list[str],
        handler: Callable[[list[str]], int],
    ) -> None:
        if self._busy_action == action_name:
            logging.info("Duplicate %s request received while action is already active.", action_name)
            return
        if self._busy_action is not None:
            logging.info("Ignoring %s request while %s is active.", action_name, self._busy_action)
            return

        self._busy_action = action_name
        try:
            handler(list(paths))
        finally:
            self._busy_action = None


def main() -> int:
    configure_logging()
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", nargs="?", help="Optional PDF to open on startup")
    args = parser.parse_args()

    app = QApplication([])
    app.setQuitOnLastWindowClosed(False)

    try:
        runtime = TrayRuntime(app)
    except RuntimeError as exc:
        logging.info("Tray startup aborted: %s", exc)
        QMessageBox.information(None, "PDF Page Editor", str(exc))
        return 0

    if args.pdf:
        logging.info("Tray started with PDF argument: %s", args.pdf)
        runtime.open_pdf(args.pdf)

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
