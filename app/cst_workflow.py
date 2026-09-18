"""Local CST export data. Facility labels stay exact in the submission manifest."""
from __future__ import annotations

import json
import shutil
import tempfile
from datetime import date, datetime
from pathlib import Path
from urllib.request import urlopen
from uuid import uuid4

from . import pdf_ops


# The list the RMRR PWA's facility field loads (facility-autocomplete.js).
FACILITIES_URL = "https://rmrrprodstorage.blob.core.windows.net/public/Facilities.txt"


def fetch_facilities(cache: Path) -> list[str]:
    """Load the deployed list; fall back to the last downloaded copy when offline."""
    try:
        with urlopen(FACILITIES_URL, timeout=10) as response:
            raw = response.read()
        names = parse_facilities(raw)
        cache.parent.mkdir(parents=True, exist_ok=True)
        cache.write_bytes(raw)
        return names
    except OSError:
        if not cache.exists():
            raise
        return parse_facilities(cache.read_bytes())


def parse_facilities(raw: bytes) -> list[str]:
    text = raw.decode("utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig")
    names = list(dict.fromkeys(line.strip() for line in text.splitlines() if line.strip()))
    if not names:
        raise ValueError("The facilities list contains no names.")
    return names


def export_cst_batch(source: Path, password: str | None, document_date: date,
                     sections: list[tuple[int, int, str, str]], facilities: list[str],
                     page_order: list[int], rotations: dict[int, int],
                     output_root: Path, *, technician: str) -> Path:
    """Publish a complete batch atomically. Sections are (start, end, facility, visit type)."""
    if not sections or not page_order:
        raise ValueError("Choose at least one section.")
    expected = 1
    for start, end, facility, _ in sections:
        if start != expected or end < start or end > len(page_order):
            raise ValueError("Sections must cover every page exactly once.")
        if facility not in facilities:
            raise ValueError("Select an exact facility name from the list for every section.")
        expected = end + 1
    if expected != len(page_order) + 1:
        raise ValueError("Sections must cover every page exactly once.")
    output_root.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix=".cst-pending-", dir=output_root))
    destination = output_root / f"CST {document_date.isoformat()} {datetime.now():%H%M%S}-{uuid4().hex[:8]}"
    try:
        names = {start: f"{facility} {document_date.isoformat()}" for start, _, facility, _ in sections}
        outputs = pdf_ops.split_pdf(source, password, [s[0] for s in sections], names,
                                    rotations, page_order, delete_source=False, output_dir=stage)
        manifest = {
            "technician": technician, "document_date": document_date.isoformat(),
            "source_path": str(source),
            "sections": [
                {"facility": facility, "visit_type": visit_type, "file": output.name,
                 "source_pages": page_order[start - 1:end]}
                for (start, end, facility, visit_type), output in zip(sections, outputs)
            ],
        }
        (stage / "batch.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
        stage.rename(destination)
        return destination
    except Exception:
        shutil.rmtree(stage)
        raise
