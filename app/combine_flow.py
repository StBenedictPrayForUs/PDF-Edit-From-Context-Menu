from __future__ import annotations

import ctypes
import logging
import os
import subprocess
import time
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QVBoxLayout,
    QWidget,
)

from . import pdf_ops


class CombineProgressWindow(QWidget):
    def __init__(self, total_files: int, title: str, initial_message: str) -> None:
        super().__init__()
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowTitle(title)
        self.setMinimumWidth(520)
        self.setMinimumHeight(260)

        layout = QVBoxLayout(self)

        self.status_label = QLabel(initial_message)
        self.status_label.setWordWrap(True)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, max(total_files, 1))
        self.progress_bar.setValue(0)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)

        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.log_view, 1)

    def update_progress(self, current: int, total: int, message: str) -> None:
        self.progress_bar.setRange(0, max(total, 1))
        self.progress_bar.setValue(min(current, max(total, 1)))
        self.status_label.setText(message)
        self.log_view.appendPlainText(message)
        logging.info("Combine progress %s/%s: %s", current, total, message)
        QApplication.processEvents()


def _open_with_default_app(path: Path) -> None:
    try:
        os.startfile(str(path))
        logging.info("Opened output with default app: %s", path)
    except Exception:
        logging.exception("Failed to open output with default app: %s", path)


def _selected_items_from_foreground_explorer() -> list[str]:
    script = r"""
Add-Type -Namespace Win32 -Name Native -MemberDefinition '[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();'
$foreground = [int64][Win32.Native]::GetForegroundWindow()
$shell = New-Object -ComObject Shell.Application
$matches = @()
foreach ($window in $shell.Windows()) {
    try {
        $items = $window.Document.SelectedItems()
        if (-not $items -or $items.Count -eq 0) {
            continue
        }
        $paths = @($items | ForEach-Object { $_.Path })
        $matches += [PSCustomObject]@{
            IsForeground = ([int64]$window.HWND -eq $foreground)
            Count = $paths.Count
            Paths = $paths
        }
    }
    catch {
    }
}
if ($matches.Count -gt 0) {
    $best = $matches |
        Sort-Object -Property @{Expression = { if ($_.IsForeground) { 1 } else { 0 } }; Descending = $true },
                              @{Expression = { $_.Count }; Descending = $true } |
        Select-Object -First 1
    $best.Paths
}
"""
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    try:
        result = subprocess.run(
            ["powershell", "-NoLogo", "-NonInteractive", "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            timeout=4,
            check=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
    except Exception:
        logging.exception("Failed to query Explorer selection.")
        return []

    if result.returncode != 0:
        logging.warning("Explorer selection query failed: %s", result.stderr.strip())
        return []

    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def _bring_widget_to_front(widget: QWidget) -> None:
    widget.show()
    widget.raise_()
    widget.activateWindow()
    QApplication.processEvents()
    try:
        ctypes.windll.user32.SetForegroundWindow(int(widget.winId()))
    except Exception:
        pass


def _prompt_save_path(title: str, suggested_path: Path) -> str:
    host = QWidget()
    host.setWindowTitle(title)
    host.setWindowFlag(Qt.Tool, True)
    host.setWindowFlag(Qt.WindowStaysOnTopHint, True)
    host.setAttribute(Qt.WA_DontShowOnScreen, False)
    host.resize(1, 1)
    _bring_widget_to_front(host)

    dialog = QFileDialog(host, title, str(suggested_path.parent), "PDF Files (*.pdf)")
    dialog.setAcceptMode(QFileDialog.AcceptSave)
    dialog.setFileMode(QFileDialog.AnyFile)
    dialog.selectFile(suggested_path.name)
    _bring_widget_to_front(dialog)
    try:
        if dialog.exec() != QFileDialog.Accepted:
            return ""
        selected_files = dialog.selectedFiles()
        return selected_files[0] if selected_files else ""
    finally:
        dialog.close()
        host.close()


def _resolve_source_paths(raw_paths: list[str], mode: str) -> list[Path]:
    use_explorer_selection = False
    filtered_paths = list(raw_paths)
    if filtered_paths and filtered_paths[0] == "--from-explorer-selection":
        use_explorer_selection = True
        filtered_paths = filtered_paths[1:]

    logging.info("%s raw paths: %s", mode, filtered_paths)
    if use_explorer_selection:
        explorer_paths: list[str] = []
        for attempt in range(5):
            explorer_paths = _selected_items_from_foreground_explorer()
            if mode == "Convert image":
                explorer_paths = [
                    path for path in explorer_paths
                    if Path(path).suffix.lower() in pdf_ops.SUPPORTED_IMAGE_EXTENSIONS
                ]
            logging.info("%s Explorer selection attempt %s: %s", mode, attempt + 1, explorer_paths)
            if explorer_paths:
                break
            time.sleep(0.15)
        if explorer_paths:
            filtered_paths = explorer_paths
        else:
            logging.warning("%s selection lookup returned no items; using raw paths.", mode)

    source_paths = pdf_ops.validate_combine_sources(filtered_paths)
    if mode == "Convert image":
        for source_path in source_paths:
            if source_path.suffix.lower() not in pdf_ops.SUPPORTED_IMAGE_EXTENSIONS:
                raise ValueError("Convert to PDF only supports image files.")
    return source_paths


def run_combine_dialog(raw_paths: list[str]) -> int:
    logging.info("Combine dialog starting. Incoming args=%s", raw_paths)
    try:
        source_paths = _resolve_source_paths(raw_paths, "Combine")
        logging.info("Combine source validation succeeded. Sources=%s", source_paths)

        app = QApplication.instance() or QApplication([])
        app.setQuitOnLastWindowClosed(False)

        suggested_path = pdf_ops.default_combined_output_path(source_paths)
        logging.info("Opening save dialog for combined output. Suggested path=%s", suggested_path)
        chosen_path = _prompt_save_path("Save Combined PDF", suggested_path)
        if not chosen_path:
            logging.info("Combine save dialog canceled by user.")
            return 0

        destination = Path(chosen_path)
        if destination.suffix.lower() != ".pdf":
            destination = destination.with_suffix(".pdf")
        logging.info("Combine save dialog returned destination=%s", destination)

        progress_window = CombineProgressWindow(
            len(source_paths),
            title="Combine to PDF",
            initial_message="Preparing combine job...",
        )
        logging.info("Showing combine progress window for %s source file(s).", len(source_paths))
        progress_window.show()
        progress_window.raise_()
        progress_window.activateWindow()
        progress_window.update_progress(0, len(source_paths), "Starting combine job...")

        def on_progress(current: int, total: int, message: str) -> None:
            progress_window.update_progress(current, total, message)

        try:
            output_path = pdf_ops.combine_documents_to_pdf(
                source_paths,
                destination,
                progress_callback=on_progress,
                delete_sources=True,
            )
        except Exception as exc:
            logging.exception("Combine to PDF failed. Sources=%s Destination=%s", source_paths, destination)
            progress_window.close()
            QMessageBox.critical(None, "Combine to PDF Failed", str(exc))
            return 1

        progress_window.update_progress(len(source_paths), len(source_paths), f"Saved {output_path.name}")
        progress_window.close()
        logging.info("Combined %s source file(s) into %s", len(source_paths), output_path)
        _open_with_default_app(output_path)
        return 0
    except Exception as exc:
        logging.exception("Combine source validation failed for %s", raw_paths)
        try:
            ctypes.windll.user32.MessageBoxW(None, str(exc), "PDF Splitter", 0x10)
        except Exception:
            pass
        return 1


def run_convert_image_dialog(raw_paths: list[str]) -> int:
    logging.info("Convert image dialog starting. Incoming args=%s", raw_paths)
    try:
        source_paths = _resolve_source_paths(raw_paths, "Convert image")
        if not source_paths:
            raise ValueError("Convert to PDF requires at least one image file.")

        app = QApplication.instance() or QApplication([])
        app.setQuitOnLastWindowClosed(False)

        if len(source_paths) == 1:
            suggested_path = source_paths[0].with_suffix(".pdf")
        else:
            suggested_path = pdf_ops.default_combined_output_path(source_paths)
        logging.info("Opening save dialog for converted image. Suggested path=%s", suggested_path)
        chosen_path = _prompt_save_path("Save PDF", suggested_path)
        if not chosen_path:
            logging.info("Convert image save dialog canceled by user.")
            return 0

        destination = Path(chosen_path)
        if destination.suffix.lower() != ".pdf":
            destination = destination.with_suffix(".pdf")
        logging.info("Convert image save dialog returned destination=%s", destination)

        progress_window = CombineProgressWindow(
            len(source_paths),
            title="Convert to PDF",
            initial_message="Preparing image conversion...",
        )
        logging.info("Showing convert progress window for %s source image(s).", len(source_paths))
        progress_window.show()
        progress_window.raise_()
        progress_window.activateWindow()
        if len(source_paths) == 1:
            start_message = f"Starting {source_paths[0].name}..."
        else:
            start_message = f"Starting {len(source_paths)} image conversion..."
        progress_window.update_progress(0, len(source_paths), start_message)

        def on_progress(current: int, total: int, message: str) -> None:
            progress_window.update_progress(current, total, message)

        try:
            output_path = pdf_ops.combine_documents_to_pdf(
                source_paths,
                destination,
                progress_callback=on_progress,
                delete_sources=True,
            )
        except Exception as exc:
            logging.exception("Convert image to PDF failed. Sources=%s Destination=%s", source_paths, destination)
            progress_window.close()
            QMessageBox.critical(None, "Convert to PDF Failed", str(exc))
            return 1

        progress_window.update_progress(len(source_paths), len(source_paths), f"Saved {output_path.name}")
        progress_window.close()
        logging.info("Converted %s image(s) into %s", len(source_paths), output_path)
        _open_with_default_app(output_path)
        return 0
    except Exception as exc:
        logging.exception("Convert image setup failed for %s", raw_paths)
        try:
            ctypes.windll.user32.MessageBoxW(None, str(exc), "PDF Splitter", 0x10)
        except Exception:
            pass
        return 1
