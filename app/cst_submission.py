"""Submit CST exports through the same public V2 API as the RMRR scanner.

Each section's submission ID is saved before sending. The server ignores a
repeated ID, so retrying an interrupted batch never duplicates an upload.
"""
from __future__ import annotations

import base64
import hashlib
import json
import re
import shutil
from datetime import date
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from uuid import uuid4

from .outlook_intake import INTAKE_DIR

API_BASE = "https://rmrr-prod-functions-eyfnceescvcmg4eg.westus2-01.azurewebsites.net"
MAX_PDF_BYTES = 10 * 1024 * 1024
RECEIVED_STATUSES = {"received", "processing_pa", "processing_reporting", "completed",
                     "retryable_failure", "reporting_retryable_failure"}


def post_section(payload: dict) -> tuple[int, dict]:
    request = Request(API_BASE + "/tracking/v2", data=json.dumps(payload).encode("utf-8"), method="POST",
                      headers={"Content-Type": "application/json", "Accept": "application/json"})
    try:
        with urlopen(request, timeout=120) as response:
            return response.status, json.loads(response.read())
    except HTTPError as exc:
        try:
            return exc.code, json.loads(exc.read())
        except ValueError:
            return exc.code, {}


def inspect_batch(manifest_path: Path) -> tuple[dict, list[dict]]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    visit_date = date.fromisoformat(manifest["document_date"])
    prepared = []
    for section in manifest["sections"]:
        path = manifest_path.parent / section["file"]
        if path.stat().st_size > MAX_PDF_BYTES:
            raise ValueError(f"{path.name} exceeds the RMRR 10 MB limit. Divide it into smaller sections.")
        # Same normalization as RMRR-PWA scanner-form/getOfficialFacilityName:
        # Facilities.txt contains Official Name/Nickname display labels.
        official = section["facility"].split("/", 1)[0].strip()
        safe_facility = re.sub(r'[/\\:*?"<>|]', "_", official)
        # Five digits like the RMRR app's suffix, which downstream tools trim. Derived
        # from the content (not the clock) so a retry sends an identical payload.
        suffix = f"{int(hashlib.sha256(path.read_bytes()).hexdigest(), 16) % 100000:05d}"
        visit_type = section["visit_type"]
        technician = manifest["technician"]
        metadata = {"technicianName": technician, "date": visit_date.isoformat(), "visitType": visit_type,
                    "sourceKind": "scanner",
                    "filename": f"{visit_type} {technician} {visit_date:%m%d%y} {safe_facility} {suffix}.pdf",
                    "facilityNames": [{"name": official, "popIn": False}], "notes": ""}
        prepared.append({"metadata": metadata, "path": path})
    if not prepared:
        raise ValueError("The batch has no sections.")
    return manifest, prepared


def valid_receipt(status: int, body: dict, submission_id: str) -> bool:
    return (status in (200, 201) and body.get("submissionId") == submission_id
            and body.get("accepted") is True and body.get("savedToServer") is True
            and not body.get("payloadMismatch") and body.get("status") in RECEIVED_STATUSES)


def submit_batch(manifest_path: Path, post=post_section, progress=lambda message: None) -> int:
    _, prepared = inspect_batch(manifest_path)
    state_path = manifest_path.parent / "submission.json"
    if state_path.exists():
        state = json.loads(state_path.read_text(encoding="utf-8"))
    else:
        state = {"jobs": [{"submission_id": str(uuid4()), "received": False} for _ in prepared]}
        state_path.write_text(json.dumps(state, indent=2), encoding="utf-8")
    if len(state["jobs"]) != len(prepared):
        raise ValueError("The batch receipt record does not match the batch.")
    for index, (job, item) in enumerate(zip(state["jobs"], prepared), 1):
        if job["received"]:
            progress(f"{index}/{len(prepared)} already received")
            continue
        progress(f"Sending {index}/{len(prepared)} — {item['metadata']['facilityNames'][0]['name']}")
        payload = {**item["metadata"], "submissionId": job["submission_id"],
                   "pdfBase64": base64.b64encode(item["path"].read_bytes()).decode("ascii")}
        status, body = post(payload)
        if not valid_receipt(status, body, job["submission_id"]):
            reason = ("the server holds different data for this submission ID"
                      if body.get("payloadMismatch") else f"HTTP {status}")
            raise RuntimeError(f"Section {index} has no confirmed receipt ({reason}).")
        job.update(received=True, receipt=body)
        state_path.write_text(json.dumps(state, indent=2), encoding="utf-8")
    return len(prepared)


def discard_batch(manifest_path: Path) -> Path:
    """After every section is received: drop the local splits and the downloaded email copy.

    Returns the source PDF's path, which identifies the intake email.
    """
    source = Path(json.loads(manifest_path.read_text(encoding="utf-8"))["source_path"])
    shutil.rmtree(manifest_path.parent)
    # Only the intake copy is disposable (the email keeps the original); a PDF
    # opened by hand from elsewhere is left alone.
    if source.parent == INTAKE_DIR:
        source.unlink(missing_ok=True)
    return source