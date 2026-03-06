from __future__ import annotations

import argparse
import logging
from pathlib import Path

from PySide6.QtCore import QObject, Signal
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import QApplication, QFileDialog, QMenu, QMessageBox, QStyle, QSystemTrayIcon

from .editor_window import PdfEditorWindow
from .ipc import IpcServer
from .logging_utils import configure_logging


class MessageBridge(QObject):
    received = Signal(dict)


class TrayRuntime:
    def __init__(self, app: QApplication) -> None:
        self.app = app
        self.window: PdfEditorWindow | None = None

        self.bridge = MessageBridge()
        self.bridge.received.connect(self._handle_ipc_message)

        self.server = IpcServer(self.bridge.received.emit)
        if not self.server.start():
            raise RuntimeError("Another PDF Splitter tray instance is already running.")
        logging.info("Tray runtime started and IPC server listening.")

        self.tray = QSystemTrayIcon(self._icon())
        self.tray.setToolTip("PDF Splitter")
        self.tray.activated.connect(self._on_tray_activated)
        self.tray.setContextMenu(self._build_menu())
        self.tray.show()

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
        menu.addAction(quit_action)
        return menu

    def shutdown(self) -> None:
        logging.info("Tray runtime shutting down.")
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
            QMessageBox.critical(None, "PDF Splitter Error", "Unexpected error while opening PDF.")
            return
        if not loaded:
            logging.warning("PDF load was canceled or failed: %s", resolved)

    def _handle_ipc_message(self, payload: dict) -> None:
        action = payload.get("action")
        if action != "open_pdf":
            return

        raw_path = payload.get("path")
        if not raw_path:
            return

        path = Path(raw_path)
        if not path.exists():
            QMessageBox.warning(None, "PDF Not Found", f"Could not find:\n{raw_path}")
            return

        self.open_pdf(str(path))

    def _on_tray_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self._show_window()


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
        QMessageBox.information(None, "PDF Splitter", str(exc))
        return 0

    if args.pdf:
        logging.info("Tray started with PDF argument: %s", args.pdf)
        runtime.open_pdf(args.pdf)

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
