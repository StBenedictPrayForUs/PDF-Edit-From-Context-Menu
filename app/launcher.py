from __future__ import annotations

import argparse
import ctypes
import logging
import subprocess
import sys
import time
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


def main() -> int:
    log_path = configure_logging()
    parser = argparse.ArgumentParser(description="Launcher entrypoint used by Windows context menu")
    parser.add_argument("pdf", nargs="?", help="Path to PDF")
    args = parser.parse_args()

    if not args.pdf:
        return 0

    source = str(Path(args.pdf).resolve())
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
