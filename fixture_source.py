"""
Offline stand-in for Outlook, used by the tests and by `--source fixtures`.

It produces exactly the same EmailMessage objects the Outlook adapter does, so
the whole pipeline - parsing, grouping, the Excel ledger, the digest HTML - can
be exercised and proved on a machine with no Outlook at all.

A fixture is a JSON file:

    {
      "subject": "PROD alert",
      "sender":  "monitor@example.com",
      "received": "2026-09-28T14:02:00",
      "folder":  "Inbox\\\\Alerts\\\\Prod",
      "body":    "Error Message: Payment gateway timeout: case ID 88421"
    }

A plain .txt file also works: the whole file is the body, the filename is the
subject, and the file's modification time is the received time.
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from message import EmailMessage, to_naive_local


def fetch_messages(directory: Path, since: dt.datetime) -> list[EmailMessage]:
    """Read fixture emails from a directory.

    Args:
        directory: Folder holding .json and/or .txt fixtures.
        since: Only fixtures received at or after this moment are returned.

    Returns:
        The matching messages, oldest first.

    Raises:
        FileNotFoundError: If the directory does not exist.
        ValueError: If a .json fixture is malformed, naming the file - a silent
            skip would make a test pass for the wrong reason.
    """
    if not directory.is_dir():
        raise FileNotFoundError(f"Fixture directory not found: {directory}")

    messages: list[EmailMessage] = []
    for path in sorted(directory.iterdir()):
        if path.suffix.lower() == ".json":
            messages.append(_from_json(path))
        elif path.suffix.lower() == ".txt":
            messages.append(_from_text(path))

    return sorted(
        (message for message in messages if message.received >= since),
        key=lambda message: message.received,
    )


def _from_json(path: Path) -> EmailMessage:
    """Build a message from a JSON fixture.

    Args:
        path: The fixture file.

    Returns:
        The message.

    Raises:
        ValueError: If the file is not valid JSON or lacks a body.
    """
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"{path.name} is not valid JSON: {exc}") from exc

    if "body" not in payload:
        raise ValueError(f"{path.name} has no 'body' field")

    received = payload.get("received")
    moment = to_naive_local(received) if received else _modified_at(path)

    return EmailMessage(
        entry_id=str(payload.get("entry_id") or path.stem),
        folder_path=str(payload.get("folder") or "Inbox\\Fixtures"),
        subject=str(payload.get("subject") or path.stem),
        sender=str(payload.get("sender") or "fixture@example.com"),
        received=moment,
        body=str(payload["body"]),
    )


def _from_text(path: Path) -> EmailMessage:
    """Build a message from a plain-text fixture.

    Args:
        path: The fixture file.

    Returns:
        The message.
    """
    return EmailMessage(
        entry_id=path.stem,
        folder_path="Inbox\\Fixtures",
        subject=path.stem,
        sender="fixture@example.com",
        received=_modified_at(path),
        body=path.read_text(encoding="utf-8"),
    )


def _modified_at(path: Path) -> dt.datetime:
    """Return a file's modification time as a naive local datetime.

    Args:
        path: The file.

    Returns:
        The modification time.
    """
    return dt.datetime.fromtimestamp(path.stat().st_mtime).replace(microsecond=0)
