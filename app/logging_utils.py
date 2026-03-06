from __future__ import annotations

import logging
from pathlib import Path


def configure_logging() -> Path:
    log_dir = Path.home() / "AppData" / "Local" / "PDFSplitter"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "app.log"

    if not logging.getLogger().handlers:
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(name)s: %(message)s",
            handlers=[logging.FileHandler(log_path, encoding="utf-8")],
        )
    return log_path
