from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings

LAST_OUTPUT_DIR_KEY = "paths/last_output_dir"


def load_last_output_dir(settings: QSettings) -> Path | None:
    raw_value = settings.value(LAST_OUTPUT_DIR_KEY, "", str)
    if not raw_value:
        return None

    path = Path(raw_value)
    if path.is_dir():
        return path
    return None


def save_last_output_dir(settings: QSettings, directory: str | Path) -> None:
    resolved_dir = Path(directory).resolve()
    settings.setValue(LAST_OUTPUT_DIR_KEY, str(resolved_dir))


def apply_last_output_dir(suggested_path: Path, last_output_dir: Path | None) -> Path:
    if last_output_dir is None:
        return suggested_path
    return last_output_dir / suggested_path.name

