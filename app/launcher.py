from __future__ import annotations

import argparse
import ctypes
import logging
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

from .ipc import send_message
from .logging_utils import configure_logging


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


def _send_action_via_ipc(action: str, paths: list[str]) -> bool:
    return send_message({"action": action, "paths": paths}, timeout=1.2)


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


def main() -> int:
    _write_startup_trace(sys.argv[1:])
    log_path = configure_logging()
    mode, paths = parse_args(sys.argv[1:])
    logging.info("Launcher started. Mode=%s Paths=%s", mode, paths)
    if mode == "combine":
        action = "combine_documents"
    elif mode == "convert-image":
        action = "convert_images"
    else:
        action = ""
    if action:
        if _send_action_via_ipc(action, paths):
            logging.info("IPC delivery succeeded for %s.", action)
            return 0

        project_root = Path(__file__).resolve().parents[1]
        logging.info("IPC delivery failed for %s; starting tray runtime.", action)
        _start_tray_runtime(project_root)
        for _ in range(20):
            time.sleep(0.25)
            if _send_action_via_ipc(action, paths):
                logging.info("IPC delivery succeeded after tray startup for %s.", action)
                return 0

        logging.error("Failed to deliver %s message after startup. Paths=%s", action, paths)
        _show_error_box(
            "PDF Splitter could not open the requested action.\n"
            f"Check log for details:\n{log_path}"
        )
        return 1

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
