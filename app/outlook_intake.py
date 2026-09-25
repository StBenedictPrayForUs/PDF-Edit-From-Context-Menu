"""Optional Outlook Classic polling. COM stays entirely on the worker thread."""
from __future__ import annotations

import hashlib
import logging
import re
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path

MAILBOX = "office@rockymountainrespiratory.com"
# Sender address -> technician name as the RMRR form files it.
SENDERS = {"jamesg@rockymountainrespiratory.com": "James",
           "ryanmcbride30377@gmail.com": "Ryan",
           "ryan@rockymountainrespiratory.com": "Ryan"}
TECHNICIANS = list(dict.fromkeys(SENDERS.values()))
SMTP_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
MESSAGE_ID_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
MAPI_E_NOT_FOUND = -2147221233
LOOKBACK = timedelta(days=14)
INTAKE_DIR = Path.home() / "AppData" / "Local" / "PDFSplitter" / "CST Intake"


def extra_body_text(body: str) -> str:
    # Only discard the observed, exact signature lines, not arbitrary quoted text.
    lines = []
    for line in body.splitlines():
        line = line.strip()
        if not line or line == "Sent from my T-Mobile 5G Device":
            continue
        if re.fullmatch(r"Get Outlook for Android(?:\s*<https://aka\.ms/AAb9ysg/?>)?", line):
            continue
        lines.append(line)
    return "\n".join(lines)


class IntakeStore:
    def __init__(self, root: Path):
        self.root = root
        root.mkdir(parents=True, exist_ok=True)
        self.database = root / "intake.sqlite"
        with self.connect() as db:
            db.execute("CREATE TABLE IF NOT EXISTS jobs (id TEXT PRIMARY KEY, path TEXT NOT NULL, note TEXT NOT NULL, done INTEGER NOT NULL DEFAULT 0)")
            columns = [row[1] for row in db.execute("PRAGMA table_info(jobs)")]
            for column, default in (("technician", "James"), ("entry_id", ""), ("store_id", "")):
                if column not in columns:
                    db.execute(f"ALTER TABLE jobs ADD COLUMN {column} TEXT NOT NULL DEFAULT '{default}'")

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.database, timeout=10)
        try:
            with db:
                yield db
        finally:
            db.close()

    def contains(self, key: str) -> bool:
        with self.connect() as db:
            return db.execute("SELECT 1 FROM jobs WHERE id=?", (key,)).fetchone() is not None

    def add(self, key: str, path: Path, note: str, technician: str, entry_id: str, store_id: str) -> None:
        with self.connect() as db:
            db.execute("INSERT OR IGNORE INTO jobs(id,path,note,technician,entry_id,store_id) VALUES (?,?,?,?,?,?)",
                       (key, str(path), note, technician, entry_id, store_id))

    def email(self, key: str) -> tuple[str, str] | None:
        with self.connect() as db:
            return db.execute("SELECT entry_id,store_id FROM jobs WHERE id=? AND entry_id<>''", (key,)).fetchone()

    def pending(self) -> list[tuple[str, str, str, str]]:
        with self.connect() as db:
            return db.execute("SELECT id,path,note,technician FROM jobs WHERE done=0 ORDER BY rowid").fetchall()

    def complete(self, key: str) -> None:
        with self.connect() as db:
            db.execute("UPDATE jobs SET done=1 WHERE id=?", (key,))

    def pending_emails(self) -> list[tuple[str, str, str]]:
        with self.connect() as db:
            return db.execute("SELECT id,entry_id,store_id FROM jobs WHERE done=0 AND entry_id<>''").fetchall()

    def dismiss(self, key: str) -> None:
        """Never process this PDF: drop it from the queue and delete the local copy."""
        with self.connect() as db:
            row = db.execute("SELECT path FROM jobs WHERE id=?", (key,)).fetchone()
            db.execute("UPDATE jobs SET done=1 WHERE id=?", (key,))
        if row:
            try:
                Path(row[0]).unlink(missing_ok=True)
            except OSError:
                logging.exception("Dismissed CST PDF could not be deleted: %s", row[0])


def sender_address(mail) -> str:
    try:
        return str(mail.PropertyAccessor.GetProperty(SMTP_PROPERTY)).lower()
    except Exception:
        if str(mail.SenderEmailType).upper() == "EX":
            user = mail.Sender.GetExchangeUser()
            return str(user.PrimarySmtpAddress).lower() if user else ""
        return str(mail.SenderEmailAddress).lower()


def left_inbox(namespace, inbox, entry_id: str, store_id: str) -> bool:
    try:
        mail = namespace.GetItemFromID(entry_id, store_id)
    except Exception as exc:
        # Deleted, or moved (Exchange gives moved items a new EntryID). Any other
        # failure, like Outlook being busy, keeps the job queued.
        info = exc.args[2] if len(exc.args) > 2 and isinstance(exc.args[2], tuple) else ()
        return MAPI_E_NOT_FOUND in (exc.args[:1] + info[5:6])
    return str(mail.Parent.EntryID) != str(inbox.EntryID)


def prune_left_inbox(namespace, store: IntakeStore, inbox) -> None:
    """Emails handled in Outlook before being processed here no longer need to open."""
    for key, entry_id, store_id in store.pending_emails():
        try:
            if left_inbox(namespace, inbox, entry_id, store_id):
                logging.info("CST email left the Inbox; dropping queued PDF %s", key)
                store.dismiss(key)
        except Exception:
            logging.exception("Could not check whether a queued CST email is still in the Inbox")


def scan_outlook(namespace, store: IntakeStore, since: datetime, stop: threading.Event) -> int:
    inbox = None
    for account in namespace.Accounts:
        if str(account.SmtpAddress).lower() == MAILBOX:
            inbox = account.DeliveryStore.GetDefaultFolder(6)
            break
    if inbox is None:
        # A shared mailbox can be mounted as a store without being an Account.
        for mailbox in namespace.Stores:
            if str(mailbox.DisplayName).lower() == MAILBOX:
                inbox = mailbox.GetDefaultFolder(6)
                break
    if inbox is None:
        raise RuntimeError("Office mailbox not found in Outlook Classic.")
    prune_left_inbox(namespace, store, inbox)

    # Outlook compares in local time. Don't re-check ReceivedTime in Python:
    # pywin32 labels that local value as UTC, which shifts it by the UTC offset.
    local_since = since.astimezone().strftime("%m/%d/%Y %I:%M %p")
    items = inbox.Items.Restrict(f"[ReceivedTime] >= '{local_since}'")
    items.Sort("[ReceivedTime]", False)
    saved = 0
    failures = 0
    for mail in items:
        if stop.is_set():
            break
        try:
            technician = SENDERS.get(sender_address(mail)) if mail.Class == 43 else None
            if technician is None:
                continue
            try:
                message_id = str(mail.PropertyAccessor.GetProperty(MESSAGE_ID_PROPERTY))
            except Exception:
                message_id = str(mail.EntryID)
            note = extra_body_text(str(mail.Body or ""))
            for index in range(1, mail.Attachments.Count + 1):
                if stop.is_set():
                    break
                attachment = mail.Attachments.Item(index)
                if Path(str(attachment.FileName)).suffix.lower() != ".pdf":
                    continue
                identity = f"{inbox.StoreID}|{message_id}|{index}"
                key = hashlib.sha256(identity.encode()).hexdigest()
                if store.contains(key):
                    continue
                # Opaque local filename: never trust attachment paths from the sender.
                path = store.root / f"{key}.pdf"
                temporary = path.with_suffix(".partial")
                attachment.SaveAsFile(str(temporary))
                if temporary.stat().st_size == 0:
                    raise ValueError("An attachment download was empty.")
                temporary.replace(path)
                store.add(key, path, note, technician, str(mail.EntryID), str(inbox.StoreID))
                saved += 1
        except Exception:
            failures += 1
            logging.exception("CST intake could not process a message; it will retry")
    if failures:
        raise RuntimeError(f"{failures} email(s) need a retry; saved attachments remain queued.")
    return saved


def file_email(store: IntakeStore, key: str) -> str:
    """Move a submitted email to its month's CST folder. Returns a note when it could not."""
    email = store.email(key)
    if email is None:
        return ""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        namespace = win32com.client.GetActiveObject("Outlook.Application").GetNamespace("MAPI")
        mail = namespace.GetItemFromID(*email)
        month = f"{mail.ReceivedTime:%m%y}"
        # e.g. 0926 > "CST 0926"; some months were named without the space.
        folders = mail.Parent.Store.GetRootFolder().Folders(month).Folders
        target = next(f for f in folders if str(f.Name).replace(" ", "").upper() == f"CST{month}")
        mail.Move(target)
        return ""
    except Exception:
        logging.exception("CST email could not be moved to its CST folder")
        return " The email could not be moved to its CST folder; it is still in the Inbox."
    finally:
        namespace = mail = folders = target = None
        pythoncom.CoUninitialize()


def poll_outlook(store: IntakeStore, enabled_iso: str, stop: threading.Event) -> str:
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return "Needs attention: install the pywin32 dependency."
    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            return "Waiting for Outlook Classic to be open."
        namespace = outlook.GetNamespace("MAPI")
        # Every scan covers the whole window; the store skips what it already has.
        since = max(datetime.fromisoformat(enabled_iso), datetime.now(timezone.utc) - LOOKBACK)
        scan_outlook(namespace, store, since, stop)
        return "Watching for CST emails • checks every 60 seconds"
    except Exception as exc:
        logging.exception("CST Outlook polling failed")
        return f"Needs attention: {exc}"
    finally:
        # Release apartment-owned proxies before uninitializing COM.
        namespace = None
        outlook = None
        pythoncom.CoUninitialize()
