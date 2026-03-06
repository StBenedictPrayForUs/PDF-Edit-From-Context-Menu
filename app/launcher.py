from __future__ import annotations

import argparse
import ctypes
import logging
import os
import subprocess
import sys
import time
from datetime import datetime
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

from .ipc import send_message
from .logging_utils import configure_logging
from . import pdf_ops

ERROR_ALREADY_EXISTS = 183


def _preferred_pythonw() -> str:
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        return str(exe)

    pythonw = exe.with_name("pythonw.exe")
    if pythonw.exists():
        return str(pythonw)
    return str(exe)


def _show_error_box(text: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(None, text, "PDF Splitter", 0x10)
    except Exception:
        pass


def _start_tray_runtime(project_root: Path, pdf_path: str | None = None) -> None:
    pythonw = _preferred_pythonw()
    cmd = [pythonw, "-m", "app.tray_runtime"]
    if pdf_path:
        cmd.append(pdf_path)
    subprocess.Popen(
        cmd,
        cwd=str(project_root),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
    )


def open_via_ipc(pdf_path: str) -> bool:
    return send_message({"action": "open_pdf", "path": pdf_path}, timeout=1.2)


def _open_with_default_app(path: Path) -> None:
    try:
        os.startfile(str(path))
        logging.info("Opened output with default app: %s", path)
    except Exception:
        logging.exception("Failed to open output with default app: %s", path)


def parse_args(argv: list[str]) -> tuple[str, list[str]]:
    if argv and argv[0].lower() == "combine":
        remaining = list(argv[1:])
        if remaining and remaining[0] == "--help":
            parser = argparse.ArgumentParser(description="Combine PDFs and images into one PDF")
            parser.add_argument("mode")
            parser.add_argument(
                "--from-explorer-selection",
                action="store_true",
                help="Read the active Explorer selection instead of trusting the raw command line.",
            )
            parser.add_argument("sources", nargs="+", help="Paths to PDFs or supported images")
            parser.parse_args(argv)

        sources: list[str] = []
        use_explorer_selection = False
        for arg in remaining:
            if arg == "--from-explorer-selection":
                use_explorer_selection = True
                continue
            sources.append(arg)

        if use_explorer_selection:
            sources.insert(0, "--from-explorer-selection")
        return "combine", sources

    if argv and argv[0].lower() == "convert-image":
        remaining = list(argv[1:])
        if remaining and remaining[0] == "--help":
            parser = argparse.ArgumentParser(description="Convert one or more images into a PDF")
            parser.add_argument("mode")
            parser.add_argument(
                "--from-explorer-selection",
                action="store_true",
                help="Read the active Explorer selection instead of trusting the raw command line.",
            )
            parser.add_argument("sources", nargs="+", help="Path(s) to image files")
            parser.parse_args(argv)

        sources: list[str] = []
        use_explorer_selection = False
        for arg in remaining:
            if arg == "--from-explorer-selection":
                use_explorer_selection = True
                continue
            sources.append(arg)

        if use_explorer_selection:
            sources.insert(0, "--from-explorer-selection")
        return "convert-image", sources

    parser = argparse.ArgumentParser(description="Launcher entrypoint used by Windows context menu")
    parser.add_argument("pdf", nargs="?", help="Path to PDF")
    args = parser.parse_args(argv)
    return "split", [args.pdf] if args.pdf else []


def _write_startup_trace(argv: list[str]) -> None:
    log_dir = Path.home() / "AppData" / "Local" / "PDFSplitter"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "app.log"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S,%f")[:-3]
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"{timestamp} INFO root: Launcher raw argv: {argv}\n")


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
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            timeout=4,
            check=False,
        )
    except Exception:
        logging.exception("Failed to query Explorer selection.")
        return []

    if result.returncode != 0:
        logging.warning("Explorer selection query failed: %s", result.stderr.strip())
        return []

    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


class _CombineLaunchMutex:
    def __init__(self, name: str) -> None:
        self._handle = ctypes.windll.kernel32.CreateMutexW(None, False, name)
        self.already_running = ctypes.GetLastError() == ERROR_ALREADY_EXISTS

    def close(self) -> None:
        if self._handle:
            ctypes.windll.kernel32.CloseHandle(self._handle)
            self._handle = None


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


def _run_combine_dialog(raw_paths: list[str]) -> int:
    mutex = _CombineLaunchMutex("Local\\PDFSplitterCombineLauncher")
    if mutex.already_running:
        logging.info("Duplicate combine launcher invocation detected; exiting early.")
        mutex.close()
        return 0

    logging.info("Combine dialog starting. Incoming args=%s", raw_paths)
    try:
        use_explorer_selection = False
        if raw_paths and raw_paths[0] == "--from-explorer-selection":
            use_explorer_selection = True
            raw_paths = raw_paths[1:]

        logging.info("Combine launcher raw paths: %s", raw_paths)
        if use_explorer_selection:
            explorer_paths: list[str] = []
            for attempt in range(5):
                explorer_paths = _selected_items_from_foreground_explorer()
                logging.info(
                    "Combine launcher Explorer selection attempt %s: %s",
                    attempt + 1,
                    explorer_paths,
                )
                if explorer_paths:
                    break
                time.sleep(0.15)
            if explorer_paths:
                raw_paths = explorer_paths
            else:
                logging.warning("Explorer selection lookup returned no items; using raw paths.")

        source_paths = pdf_ops.validate_combine_sources(raw_paths)
        logging.info("Combine source validation succeeded. Sources=%s", source_paths)

        app = QApplication.instance() or QApplication([])
        app.setQuitOnLastWindowClosed(False)

        suggested_path = pdf_ops.default_combined_output_path(source_paths)
        logging.info("Opening save dialog for combined output. Suggested path=%s", suggested_path)
        chosen_path, _ = QFileDialog.getSaveFileName(
            None,
            "Save Combined PDF",
            str(suggested_path),
            "PDF Files (*.pdf)",
        )
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
        _show_error_box(str(exc))
        return 1
    finally:
        mutex.close()


def _run_convert_image_dialog(raw_paths: list[str]) -> int:
    mutex = _CombineLaunchMutex("Local\\PDFSplitterConvertImageLauncher")
    if mutex.already_running:
        logging.info("Duplicate convert-image launcher invocation detected; exiting early.")
        mutex.close()
        return 0

    logging.info("Convert image dialog starting. Incoming args=%s", raw_paths)
    try:
        use_explorer_selection = False
        if raw_paths and raw_paths[0] == "--from-explorer-selection":
            use_explorer_selection = True
            raw_paths = raw_paths[1:]

        logging.info("Convert image raw paths: %s", raw_paths)
        if use_explorer_selection:
            explorer_paths: list[str] = []
            for attempt in range(5):
                explorer_paths = _selected_items_from_foreground_explorer()
                explorer_paths = [
                    path for path in explorer_paths
                    if Path(path).suffix.lower() in pdf_ops.SUPPORTED_IMAGE_EXTENSIONS
                ]
                logging.info(
                    "Convert image Explorer selection attempt %s: %s",
                    attempt + 1,
                    explorer_paths,
                )
                if explorer_paths:
                    break
                time.sleep(0.15)
            if explorer_paths:
                raw_paths = explorer_paths
            else:
                logging.warning("Convert image selection lookup returned no items; using raw paths.")

        source_paths = pdf_ops.validate_combine_sources(raw_paths)
        if not source_paths:
            raise ValueError("Convert to PDF requires at least one image file.")
        for source_path in source_paths:
            if source_path.suffix.lower() not in pdf_ops.SUPPORTED_IMAGE_EXTENSIONS:
                raise ValueError("Convert to PDF only supports image files.")

        app = QApplication.instance() or QApplication([])
        app.setQuitOnLastWindowClosed(False)

        if len(source_paths) == 1:
            suggested_path = source_paths[0].with_suffix(".pdf")
        else:
            suggested_path = pdf_ops.default_combined_output_path(source_paths)
        logging.info("Opening save dialog for converted image. Suggested path=%s", suggested_path)
        chosen_path, _ = QFileDialog.getSaveFileName(
            None,
            "Save PDF",
            str(suggested_path),
            "PDF Files (*.pdf)",
        )
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
        _show_error_box(str(exc))
        return 1
    finally:
        mutex.close()


def main() -> int:
    _write_startup_trace(sys.argv[1:])
    log_path = configure_logging()
    mode, paths = parse_args(sys.argv[1:])
    logging.info("Launcher started. Mode=%s Paths=%s", mode, paths)
    if mode == "combine":
        return _run_combine_dialog(paths)
    if mode == "convert-image":
        return _run_convert_image_dialog(paths)

    if not paths:
        return 0

    source = str(Path(paths[0]).resolve())
    logging.info("Launcher invoked for %s", source)

    if open_via_ipc(source):
        logging.info("IPC delivery succeeded (existing tray instance).")
        return 0

    project_root = Path(__file__).resolve().parents[1]
    logging.info("IPC delivery failed; starting tray runtime.")
    _start_tray_runtime(project_root, source)

    for _ in range(20):
        time.sleep(0.25)
        if open_via_ipc(source):
            logging.info("IPC delivery succeeded after tray startup.")
            return 0

    logging.error("Failed to deliver open_pdf message after startup. Source=%s", source)
    _show_error_box(
        "PDF Splitter could not open the file.\n"
        f"Check log for details:\n{log_path}"
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
