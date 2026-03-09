from __future__ import annotations

from dataclasses import dataclass
import logging
from pathlib import Path
import math
import re
from typing import Callable, Iterable

import fitz
from pypdf import PdfReader, PdfWriter
from PySide6.QtGui import QImage

WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

SUPPORTED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
SUPPORTED_COMBINE_EXTENSIONS = {".pdf", *SUPPORTED_IMAGE_EXTENSIONS}
BALANCED_MAX_IMAGE_DIMENSION = 2200
BALANCED_JPEG_QUALITY = 78


class PasswordRequiredError(Exception):
    pass


class InvalidPasswordError(Exception):
    pass


class UnsupportedSourceError(Exception):
    pass


@dataclass
class PdfMetadata:
    source_path: Path
    page_count: int
    base_name: str
    password: str | None


@dataclass(frozen=True)
class CompressionProfile:
    max_dimension: int
    jpeg_quality: int


BALANCED_COMPRESSION = CompressionProfile(
    max_dimension=BALANCED_MAX_IMAGE_DIMENSION,
    jpeg_quality=BALANCED_JPEG_QUALITY,
)

ProgressCallback = Callable[[int, int, str], None]


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
    if page_count <= 0:
        return []

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


def normalize_page_order(page_count: int, page_order: Iterable[int] | None = None) -> list[int]:
    if page_count < 0:
        raise ValueError("Page count cannot be negative.")

    if page_order is None:
        return list(range(1, page_count + 1))

    ordered = [int(page) for page in page_order]
    if len(ordered) != page_count:
        raise ValueError("Page order must include every page exactly once.")

    expected = set(range(1, page_count + 1))
    if set(ordered) != expected:
        raise ValueError("Page order must include every page exactly once.")

    return ordered


def sanitize_filename(name: str, fallback: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", (name or "").strip())
    clean = clean.rstrip(" .")
    if not clean:
        clean = fallback
    if clean.split(".")[0].upper() in WINDOWS_RESERVED_NAMES:
        clean = f"{clean}_file"
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


def validate_combine_sources(source_paths: Iterable[str | Path]) -> list[Path]:
    resolved_paths: list[Path] = []
    for raw_path in source_paths:
        path = Path(raw_path).resolve()
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        if path.suffix.lower() not in SUPPORTED_COMBINE_EXTENSIONS:
            raise UnsupportedSourceError(f"Unsupported file type: {path.name}")
        resolved_paths.append(path)

    if not resolved_paths:
        raise ValueError("Select at least one PDF or image to combine.")

    logging.info("Validated combine sources: %s", resolved_paths)
    return resolved_paths


def default_combined_output_path(
    source_paths: Iterable[str | Path],
    suggested_name: str = "combined.pdf",
) -> Path:
    resolved_paths = validate_combine_sources(source_paths)
    default_dir = _default_output_directory(resolved_paths)
    filename = sanitize_filename(suggested_name, "combined.pdf")
    if not filename.lower().endswith(".pdf"):
        filename = f"{filename}.pdf"
    return default_dir / filename


def _default_output_directory(source_paths: list[Path]) -> Path:
    first_parent = source_paths[0].parent
    if all(path.parent == first_parent for path in source_paths):
        return first_parent
    return first_parent


def _compression_profile(name: str) -> CompressionProfile:
    if name != "balanced":
        raise ValueError(f"Unsupported compression profile: {name}")
    return BALANCED_COMPRESSION


def _append_pdf(target: fitz.Document, source_path: Path) -> None:
    logging.info("Appending PDF source: %s", source_path)
    doc = fitz.open(str(source_path))
    try:
        target.insert_pdf(doc)
    finally:
        doc.close()


def _append_image(target: fitz.Document, source_path: Path, profile: CompressionProfile) -> None:
    logging.info("Appending image source: %s", source_path)
    image_doc = fitz.open(str(source_path))
    try:
        keep_original = source_path.suffix.lower() in {".jpg", ".jpeg"} and _can_keep_original_image(
            image_doc,
            profile,
        )
        if keep_original:
            logging.info("Keeping original image without recompression: %s", source_path)
            converted = fitz.open("pdf", image_doc.convert_to_pdf())
            try:
                target.insert_pdf(converted)
            finally:
                converted.close()
            return

        for page_index in range(image_doc.page_count):
            page = image_doc.load_page(page_index)
            pixmap = _compressed_pixmap_for_page(page, profile)
            logging.info(
                "Compressed image page %s from %s to %sx%s",
                page_index + 1,
                source_path.name,
                pixmap.width,
                pixmap.height,
            )
            page_width = max(float(pixmap.width), 1.0)
            page_height = max(float(pixmap.height), 1.0)
            output_page = target.new_page(width=page_width, height=page_height)
            output_page.insert_image(
                output_page.rect,
                stream=pixmap.tobytes("jpeg", jpg_quality=profile.jpeg_quality),
            )
    finally:
        image_doc.close()


def _can_keep_original_image(image_doc: fitz.Document, profile: CompressionProfile) -> bool:
    for page_index in range(image_doc.page_count):
        rect = image_doc.load_page(page_index).rect
        if max(rect.width, rect.height) > profile.max_dimension:
            return False
    return True


def _compressed_pixmap_for_page(page: fitz.Page, profile: CompressionProfile) -> fitz.Pixmap:
    scale = min(1.0, profile.max_dimension / max(page.rect.width, page.rect.height, 1.0))
    if math.isclose(scale, 1.0):
        matrix = fitz.Matrix(1, 1)
    else:
        matrix = fitz.Matrix(scale, scale)
    return page.get_pixmap(matrix=matrix, alpha=False)


def combine_documents_to_pdf(
    source_paths: Iterable[str | Path],
    destination_path: str | Path,
    compression_profile: str = "balanced",
    progress_callback: ProgressCallback | None = None,
    delete_sources: bool = False,
) -> Path:
    sources = validate_combine_sources(source_paths)
    destination = Path(destination_path).resolve()
    if destination.suffix.lower() != ".pdf":
        destination = destination.with_suffix(".pdf")
    destination.parent.mkdir(parents=True, exist_ok=True)

    profile = _compression_profile(compression_profile)
    logging.info(
        "Starting document combine. Source count=%s Destination=%s Profile=%s",
        len(sources),
        destination,
        compression_profile,
    )
    merged = fitz.open()
    try:
        total_steps = len(sources)
        for index, path in enumerate(sources, start=1):
            if progress_callback is not None:
                progress_callback(index - 1, total_steps, f"Processing {path.name}")
            if path.suffix.lower() == ".pdf":
                _append_pdf(merged, path)
            else:
                _append_image(merged, path, profile)

        if merged.page_count == 0:
            raise ValueError("The selected files did not produce any PDF pages.")

        if progress_callback is not None:
            progress_callback(total_steps, total_steps, f"Saving {destination.name}")
        merged.save(
            str(destination),
            garbage=3,
            deflate=True,
            deflate_images=True,
            use_objstms=True,
        )
    finally:
        merged.close()

    if delete_sources:
        for index, path in enumerate(sources, start=1):
            if progress_callback is not None:
                progress_callback(index, len(sources), f"Deleting {path.name}")
            logging.info("Deleting source after successful combine: %s", path)
            path.unlink()

    logging.info("Document combine completed successfully: %s", destination)
    return destination


def split_pdf(
    source_path: str | Path,
    password: str | None,
    split_starts: Iterable[int],
    section_names: dict[int, str],
    page_rotations: dict[int, int],
    page_order: Iterable[int] | None = None,
    delete_source: bool = False,
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
    ordered_pages = normalize_page_order(page_count, page_order)
    sections = compute_sections(len(ordered_pages), split_starts)
    written_paths: list[Path] = []

    for idx, (start, end) in enumerate(sections, start=1):
        fallback_name = f"{source.stem}_part_{idx:02d}"
        section_name = sanitize_filename(section_names.get(start, ""), fallback_name)
        destination = ensure_unique_path(out_dir / f"{section_name}.pdf")

        writer = PdfWriter()
        for page_position in range(start, end + 1):
            source_page_num = ordered_pages[page_position - 1]
            page = reader.pages[source_page_num - 1]
            rotate = int(page_rotations.get(source_page_num, 0)) % 360
            if rotate:
                page.rotate(rotate)
            writer.add_page(page)

        with destination.open("wb") as f:
            writer.write(f)

        written_paths.append(destination)

    if delete_source:
        logging.info("Deleting split source after successful export: %s", source)
        source.unlink()

    return written_paths
