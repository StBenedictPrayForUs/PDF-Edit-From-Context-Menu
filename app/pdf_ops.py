from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Iterable

import fitz
from pypdf import PdfReader, PdfWriter
from PySide6.QtGui import QImage


class PasswordRequiredError(Exception):
    pass


class InvalidPasswordError(Exception):
    pass


@dataclass
class PdfMetadata:
    source_path: Path
    page_count: int
    base_name: str
    password: str | None


def load_pdf_metadata(path: str | Path, password: str | None = None) -> PdfMetadata:
    source_path = Path(path).resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"File not found: {source_path}")

    doc = fitz.open(str(source_path))
    try:
        if doc.needs_pass:
            if password is None:
                raise PasswordRequiredError("PDF is password protected.")
            if not doc.authenticate(password):
                raise InvalidPasswordError("Invalid password.")
        page_count = doc.page_count
    finally:
        doc.close()

    return PdfMetadata(
        source_path=source_path,
        page_count=page_count,
        base_name=source_path.stem,
        password=password,
    )


def render_page_thumbnail(
    path: str | Path,
    page_index: int,
    password: str | None = None,
    max_width: int = 220,
) -> QImage:
    doc = fitz.open(str(path))
    try:
        if doc.needs_pass:
            if not password:
                raise PasswordRequiredError("PDF is password protected.")
            if not doc.authenticate(password):
                raise InvalidPasswordError("Invalid password.")

        page = doc.load_page(page_index)
        rect = page.rect
        zoom = max_width / max(rect.width, 1)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        fmt = QImage.Format_RGB888
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        return image.copy()
    finally:
        doc.close()


def compute_sections(page_count: int, split_starts: Iterable[int]) -> list[tuple[int, int]]:
    starts = {1}
    for page in split_starts:
        if 1 <= page <= page_count:
            starts.add(page)

    ordered = sorted(starts)
    sections: list[tuple[int, int]] = []
    for idx, start in enumerate(ordered):
        end = ordered[idx + 1] - 1 if idx + 1 < len(ordered) else page_count
        sections.append((start, end))
    return sections


def sanitize_filename(name: str, fallback: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", (name or "").strip())
    clean = clean.rstrip(" .")
    if not clean:
        clean = fallback
    return clean[:180]


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 1
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def split_pdf(
    source_path: str | Path,
    password: str | None,
    split_starts: Iterable[int],
    section_names: dict[int, str],
    page_rotations: dict[int, int],
    output_dir: str | Path | None = None,
) -> list[Path]:
    source = Path(source_path).resolve()
    out_dir = Path(output_dir).resolve() if output_dir else source.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(source))
    if reader.is_encrypted:
        if password is None:
            raise PasswordRequiredError("PDF is password protected.")
        if reader.decrypt(password) == 0:
            raise InvalidPasswordError("Invalid password.")

    page_count = len(reader.pages)
    sections = compute_sections(page_count, split_starts)
    written_paths: list[Path] = []

    for idx, (start, end) in enumerate(sections, start=1):
        fallback_name = f"{source.stem}_part_{idx:02d}"
        section_name = sanitize_filename(section_names.get(start, ""), fallback_name)
        destination = ensure_unique_path(out_dir / f"{section_name}.pdf")

        writer = PdfWriter()
        for page_num in range(start, end + 1):
            page = reader.pages[page_num - 1]
            rotate = int(page_rotations.get(page_num, 0)) % 360
            if rotate:
                page.rotate(rotate)
            writer.add_page(page)

        with destination.open("wb") as f:
            writer.write(f)

        written_paths.append(destination)

    return written_paths
