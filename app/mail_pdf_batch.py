from __future__ import annotations

import argparse
import hashlib
import logging
import shutil
from dataclasses import dataclass
from pathlib import Path

from .logging_utils import configure_logging
from . import pdf_ops


MANIFEST_FIELD_COUNT = 3
RESULT_SAVED = "saved"
RESULT_DUPLICATE = "duplicate"
RESULT_EMPTY = "empty"
RESULT_FAILED = "failed"


@dataclass
class BatchJob:
    job_id: str
    output_path: Path
    source_dir: Path


@dataclass
class BatchResult:
    job_id: str
    status: str
    detail: str


def _read_manifest(manifest_path: Path) -> list[BatchJob]:
    jobs: list[BatchJob] = []
    for line_number, raw_line in enumerate(
        manifest_path.read_text(encoding="utf-8-sig").splitlines(), start=1
    ):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        fields = line.split("\t")
        if len(fields) != MANIFEST_FIELD_COUNT:
            raise ValueError(
                f"Manifest line {line_number} has {len(fields)} fields, expected {MANIFEST_FIELD_COUNT}"
            )
        job_id, output_path, source_dir = fields
        jobs.append(BatchJob(job_id, Path(output_path), Path(source_dir)))
    return jobs


def _source_paths(source_dir: Path) -> list[Path]:
    if not source_dir.is_dir():
        return []
    return [
        path
        for path in sorted(source_dir.iterdir(), key=lambda item: item.name.lower())
        if path.is_file() and path.suffix.lower() in pdf_ops.SUPPORTED_COMBINE_EXTENSIONS
    ]


def _file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def content_fingerprint(source_paths: list[Path]) -> str:
    """Fingerprint the attachment set itself, not the produced PDF.

    PyMuPDF output is not byte-stable, so hashing the combined PDF would miss
    genuine duplicates. Hashing the sorted multiset of source digests means a
    resubmitted mail still matches even if the attachment names or order differ.
    """
    digests = sorted(_file_digest(path) for path in source_paths)
    return hashlib.sha256("\n".join(digests).encode("ascii")).hexdigest()


def _unique_output_path(destination: Path) -> Path:
    if not destination.exists():
        return destination
    stem = destination.stem
    suffix = destination.suffix or ".pdf"
    index = 2
    while True:
        candidate = destination.with_name(f"{stem} ({index}){suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def run_batch(
    jobs: list[BatchJob],
    compression_profile: str = "outlook-attachment",
    cleanup: bool = False,
) -> list[BatchResult]:
    results: list[BatchResult] = []
    seen: dict[str, str] = {}

    for job in jobs:
        sources = _source_paths(job.source_dir)
        if not sources:
            logging.warning("Job %s has no supported attachments in %s", job.job_id, job.source_dir)
            results.append(BatchResult(job.job_id, RESULT_EMPTY, ""))
            continue

        fingerprint = content_fingerprint(sources)
        if fingerprint in seen:
            logging.info("Job %s duplicates %s; skipping combine", job.job_id, seen[fingerprint])
            results.append(BatchResult(job.job_id, RESULT_DUPLICATE, seen[fingerprint]))
            if cleanup:
                shutil.rmtree(job.source_dir, ignore_errors=True)
            continue

        destination = _unique_output_path(job.output_path)
        try:
            output = pdf_ops.combine_documents_to_pdf(
                sources,
                destination,
                compression_profile=compression_profile,
                delete_sources=False,
            )
        except Exception as error:
            logging.exception("Job %s failed to combine", job.job_id)
            results.append(BatchResult(job.job_id, RESULT_FAILED, str(error)))
            continue

        seen[fingerprint] = str(output)
        results.append(BatchResult(job.job_id, RESULT_SAVED, str(output)))
        if cleanup:
            shutil.rmtree(job.source_dir, ignore_errors=True)

    return results


def _write_results(results_path: Path, results: list[BatchResult]) -> None:
    lines = [f"{item.job_id}\t{item.status}\t{item.detail}" for item in results]
    results_path.parent.mkdir(parents=True, exist_ok=True)
    results_path.write_text("\n".join(lines), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Collate many mails' attachments into one PDF per mail, skipping duplicate content."
    )
    parser.add_argument("--manifest", required=True, help="Tab-delimited jobId/outputPath/sourceDir file.")
    parser.add_argument("--results", required=True, help="Destination for the tab-delimited result file.")
    parser.add_argument("--profile", default="outlook-attachment", help="Compression profile name.")
    parser.add_argument("--cleanup", action="store_true", help="Delete each source folder after processing.")
    args = parser.parse_args(argv)

    configure_logging()
    try:
        jobs = _read_manifest(Path(args.manifest))
    except Exception:
        logging.exception("Failed to read manifest: %s", args.manifest)
        return 1

    logging.info("Starting mail PDF batch. Jobs=%s Profile=%s", len(jobs), args.profile)
    results = run_batch(jobs, compression_profile=args.profile, cleanup=args.cleanup)

    try:
        _write_results(Path(args.results), results)
    except Exception:
        logging.exception("Failed to write results: %s", args.results)
        return 1

    failed = sum(1 for item in results if item.status == RESULT_FAILED)
    saved = sum(1 for item in results if item.status == RESULT_SAVED)
    duplicates = sum(1 for item in results if item.status == RESULT_DUPLICATE)
    logging.info(
        "Mail PDF batch finished. Saved=%s Duplicates=%s Failed=%s", saved, duplicates, failed
    )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
