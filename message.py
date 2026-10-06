"""
Shared email record.

Both the live Outlook adapter (outlook_win.py) and the offline fixture reader
(fixture_source.py) produce this exact shape, so the rest of the program never
knows or cares which one it is talking to.  That is what makes the whole
pipeline testable on a machine that has no Outlook.
"""

from __future__ import annotations

import datetime as _datetime
from dataclasses import dataclass


@dataclass(frozen=True)
class EmailMessage:
    """One email, reduced to the only six things this program reads.

    Attributes:
        entry_id: Stable unique identifier used to avoid processing the same
            email twice.  From Outlook this is the item's EntryID.
        folder_path: Human-readable folder the message was found in, e.g.
            "Inbox\\Alerts\\Prod".  Recorded so the ledger can show where an
            error arrives from.
        subject: The email subject line.
        sender: Display name or address of the sender.
        received: Naive local datetime the message was received.
        body: Plain-text body.  HTML bodies are not used - the plain-text
            alternative is what gets parsed, because HTML markup would corrupt
            the error signatures.
    """

    entry_id: str
    folder_path: str
    subject: str
    sender: str
    received: _datetime.datetime
    body: str


def to_naive_local(value: object) -> _datetime.datetime:
    """Convert any datetime-like value to a naive local datetime.

    Outlook (via pywin32) hands back timezone-aware pywintypes datetimes, while
    fixtures and the Excel workbook carry naive ones.  Mixing the two raises
    "can't compare offset-naive and offset-aware datetimes" at the exact moment
    a comparison matters, so every datetime entering this program passes
    through here first.

    Args:
        value: A datetime, a date, or an ISO-8601 string.

    Returns:
        A naive datetime in local time.

    Raises:
        TypeError: If the value cannot be interpreted as a datetime.
    """
    if isinstance(value, _datetime.datetime):
        moment = value
    elif isinstance(value, _datetime.date):
        return _datetime.datetime(value.year, value.month, value.day)
    elif isinstance(value, str):
        text = value.strip().replace("Z", "+00:00")
        try:
            moment = _datetime.datetime.fromisoformat(text)
        except ValueError as exc:
            raise TypeError(f"Not a datetime: {value!r}") from exc
    else:
        raise TypeError(f"Not a datetime: {value!r} ({type(value).__name__})")

    if moment.tzinfo is not None:
        moment = moment.astimezone().replace(tzinfo=None)
    return moment
