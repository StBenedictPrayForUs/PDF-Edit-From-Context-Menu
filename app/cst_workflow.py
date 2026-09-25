"""Local CST export data. Facility labels stay exact in the submission manifest."""
from __future__ import annotations

import json
import re
import shutil
import tempfile
from datetime import date, datetime
from difflib import SequenceMatcher
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


SUBJECT_DATE = re.compile(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b")
STOPWORDS = {"the", "at", "of", "and", "a"}


def date_from_subject(subject: str, today: date | None = None) -> date | None:
    """'Brighton 9.2.26' or 'Cherrelyn 9-04-2026 .pdf' -> that date, if it is recent."""
    match = SUBJECT_DATE.search(subject)
    if not match:
        return None
    month, day, year = (int(part) for part in match.groups())
    try:
        found = date(year if year >= 1000 else 2000 + year % 100, month, day)
    except ValueError:
        return None
    return found if abs((found - (today or date.today())).days) <= 60 else None


def facility_from_subject(subject: str, facilities: list[str]) -> str | None:
    """Match the words before the date ('Boulder Post sleep stdy 9.24.26') to one facility.

    The longest leading phrase that resembles a run of words in a facility name wins;
    an ambiguous phrase ('Lowry') fills nothing.
    """
    words = re.findall(r"[a-z0-9]+", SUBJECT_DATE.split(subject.lower().replace("'", ""))[0])
    spans = []  # (facility, a run of consecutive words from one of its "/" names, joined)
    for facility in facilities:
        for alias in facility.lower().replace("'", "").split("/"):
            tokens = re.findall(r"[a-z0-9]+", alias)
            spans += [(facility, "".join(tokens[i:j])) for i in range(len(tokens)) for j in range(i + 1, len(tokens) + 1)]
    for size in range(len(words), 0, -1):
        phrase = words[:size]
        joined = "".join(phrase)
        if all(word in STOPWORDS for word in phrase) or len(joined) < 4:
            continue
        # Joined text tolerates spacing and plurals: 'Cityscape' ~ 'City Scape', 'Wellspring' ~ 'Wellsprings'.
        matches = {facility for facility, span in spans if SequenceMatcher(None, joined, span).ratio() >= 0.9}
        if matches:
            return matches.pop() if len(matches) == 1 else None
    return None


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
