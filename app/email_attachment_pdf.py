from __future__ import annotations

import argparse
import logging
import os
import shutil
from pathlib import Path

from .logging_utils import configure_logging
from . import pdf_ops


def _open_with_default_app(path: Path) -> None:
    try:
        os.startfile(str(path))
        logging.info("Opened output with default app: %s", path)
    except Exception:
        logging.exception("Failed to open output with default app: %s", path)


def _source_paths_from_folder(source_dir: Path) -> list[Path]:
    if not source_dir.exists():
        raise FileNotFoundError(f"Source folder not found: {source_dir}")
    if not source_dir.is_dir():
        raise NotADirectoryError(f"Source path is not a folder: {source_dir}")

    sources = [
        path
        for path in sorted(source_dir.iterdir(), key=lambda item: item.name.lower())
        if path.is_file() and path.suffix.lower() in pdf_ops.SUPPORTED_COMBINE_EXTENSIONS
    ]
    if not sources:
        raise ValueError("No supported PDF or image attachments were found.")
    return sources


def combine_email_attachments(source_dir: str | Path, output_path: str | Path, cleanup: bool = True) -> Path:
    source_folder = Path(source_dir).resolve()
    destination = Path(output_path).resolve()
    sources = _source_paths_from_folder(source_folder)

    logging.info(
        "Combining email attachments. Source folder=%s Source count=%s Output=%s",
        source_folder,
        len(sources),
        destination,
    )
    output = pdf_ops.combine_documents_to_pdf(
        sources,
        destination,
        compression_profile="outlook-attachment",
        delete_sources=False,
    )
    _open_with_default_app(output)

    if cleanup:
        try:
            shutil.rmtree(source_folder)
            logging.info("Deleted temporary attachment folder: %s", source_folder)
        except Exception:
            logging.exception("Failed to delete temporary attachment folder: %s", source_folder)

    return output


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Combine saved Outlook attachments into one PDF.")
    parser.add_argument("--source-dir", required=True, help="Folder containing saved attachments.")
    parser.add_argument("--output", required=True, help="Destination PDF path.")
    parser.add_argument("--keep-temp", action="store_true", help="Do not delete the source folder after combining.")
    args = parser.parse_args(argv)

    configure_logging()
    try:
        output = combine_email_attachments(args.source_dir, args.output, cleanup=not args.keep_temp)
    except Exception:
        logging.exception(
            "Failed to combine Outlook attachments. Source folder=%s Output=%s",
            args.source_dir,
            args.output,
        )
        return 1

    logging.info("Outlook attachment PDF saved: %s", output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
