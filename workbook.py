"""
The Excel workbook: configuration and permanent ledger in one file.

Four sheets:

    Config        key/value settings - folders, threshold, recipients, last run
    Keywords      the error-start markers, each with an Enabled flag
    Errors        one row per distinct error - the consolidated ledger
    Occurrences   one row per error instance - the history and the dedup key

Three rules hold this together:

  1. Columns owned by the person (Status, Owner, Notes) are read but NEVER
     written by the script, so a run cannot undo their triage.
  2. First Found is written once and never changed, so "when did we first see
     this" survives forever.
  3. Everything is matched on Entry ID, so running twice over the same window
     cannot double-count.
"""

from __future__ import annotations

import datetime as dt
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from message import to_naive_local
from parsing import Keyword

CONFIG_SHEET = "Config"
KEYWORDS_SHEET = "Keywords"
ERRORS_SHEET = "Errors"
OCCURRENCES_SHEET = "Occurrences"

# (key, default value, explanation shown in the sheet)
DEFAULT_CONFIG: tuple[tuple[str, str, str], ...] = (
    ("folders", "Inbox\\Alerts",
     "Outlook folders to scan, ONE PER LINE. Use Alt+Enter in Excel for a new line."),
    ("fuzzy_threshold", "0.90",
     "0 to 1. Two errors group when they are at least this similar. Lower groups more."),
    ("last_run", "",
     "Written by the script. Emails received after this are processed. Clear it to start over."),
    ("initial_lookback_days", "7",
     "Used only when last_run is empty - how far back the very first run reaches."),
    ("strip_reply_chain", "yes",
     "yes = ignore quoted history in forwarded or replied alerts."),
    ("max_errors_per_email", "0",
     "0 = count every marker found. 1 = only the first error in each email."),
    ("id_labels",
     "case, ticket, incident, request, correlation, transaction, job, order, "
     "reference, ref, session, batch, trace, record, invoice, claim, message, event",
     "Words before an identifier. These identifiers are stripped out so the error groups."),
    ("draft_to", "", "Recipients of the digest draft, separated by ';'."),
    ("draft_cc", "", "Cc of the digest draft, separated by ';'."),
    ("draft_subject", "Error Digest {date} - {new} new, {recurring} recurring",
     "Subject template. {date} {new} {recurring} {total} are filled in."),
    ("occurrence_retention_days", "180",
     "Occurrence rows older than this are removed. The Errors ledger is never pruned."),
)

DEFAULT_KEYWORDS: tuple[tuple[str, str, str], ...] = (
    ("Error Message", "Y", "The most specific marker - tried before plain 'Error'."),
    ("Exception", "Y", ""),
    ("Error", "Y", ""),
    ("Failure", "Y", ""),
    ("Failed", "Y", ""),
    ("FATAL", "Y", ""),
    ("Warning", "N", "Disabled by default - set to Y to include warnings."),
)

# Machine columns are rewritten every run.  User columns are never touched.
ERROR_HEADERS: tuple[str, ...] = (
    "Signature",        # machine
    "Sample Message",   # machine
    "Total Count",      # machine
    "Count This Run",   # machine
    "First Found",      # machine, written once
    "Last Seen",        # machine
    "Status",           # USER
    "Owner",            # USER
    "Notes",            # USER
    "Case IDs Seen",    # machine
    "Source Folders",   # machine
    "Last Updated",     # machine
)
USER_OWNED_COLUMNS: frozenset[str] = frozenset({"Status", "Owner", "Notes"})

OCCURRENCE_HEADERS: tuple[str, ...] = (
    "Received", "Folder", "Subject", "Sender",
    "Signature", "Raw Error Text", "Captured IDs", "Entry ID", "Run ID",
)

_HEADER_FILL = PatternFill("solid", fgColor="1F3864")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_USER_FILL = PatternFill("solid", fgColor="FFF2CC")
_DATE_FORMAT = "yyyy-mm-dd hh:mm"
_MAX_CASE_IDS = 50


@dataclass
class ErrorRow:
    """One row of the Errors ledger, as read from the workbook.

    Attributes:
        signature: The normalized grouping key.
        sample: Verbatim text of the first occurrence ever seen.
        total_count: All-time number of occurrences.
        first_found: When this error was first seen.
        last_seen: When it was most recently seen.
        status: The person's triage state.
        case_ids: Identifiers recorded against this error.
        folders: Folders this error has arrived in.
        excel_row: 1-based row number in the sheet.
    """

    signature: str
    sample: str
    total_count: int
    first_found: dt.datetime | None
    last_seen: dt.datetime | None
    status: str
    case_ids: list[str]
    folders: list[str]
    excel_row: int


class DigestWorkbook:
    """Reads, updates and writes the error-digest workbook.

    Open it with :meth:`open`, call :meth:`begin_run`, record errors with
    :meth:`record`, then :meth:`save`.
    """

    def __init__(self, path: Path, workbook: Workbook) -> None:
        """Store an already-loaded workbook.

        Use :meth:`open` rather than calling this directly.

        Args:
            path: Where the workbook lives on disk.
            workbook: The loaded openpyxl workbook.
        """
        self.path = path
        self.workbook = workbook
        self._errors: dict[str, ErrorRow] = {}
        self._new_signatures: set[str] = set()
        self._run_counts: dict[str, int] = {}
        self._pending_occurrences: list[tuple[Any, ...]] = []
        self._seen_entry_ids: set[str] = set()
        self._load_errors()
        self._load_entry_ids()

    # -- construction ------------------------------------------------------

    @classmethod
    def open(cls, path: Path) -> "DigestWorkbook":
        """Load the workbook, creating it with defaults if it does not exist.

        Args:
            path: Workbook path.

        Returns:
            A ready-to-use DigestWorkbook.

        Raises:
            OSError: If the file exists but cannot be read.
        """
        if path.exists():
            workbook = load_workbook(path)
            _ensure_sheets(workbook)
        else:
            workbook = _build_new_workbook()
            path.parent.mkdir(parents=True, exist_ok=True)
            workbook.save(path)
        return cls(path, workbook)

    # -- configuration -----------------------------------------------------

    @property
    def config(self) -> dict[str, str]:
        """Return every Config key/value pair as strings.

        Returns:
            A dict of setting name to raw string value.
        """
        sheet = self.workbook[CONFIG_SHEET]
        values: dict[str, str] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            key = str(row[0]).strip()
            raw = row[1] if len(row) > 1 else None
            values[key] = "" if raw is None else str(raw).strip()
        return values

    def get_str(self, key: str, default: str = "") -> str:
        """Read a Config value as text.

        Args:
            key: Setting name.
            default: Returned when the setting is missing or blank.

        Returns:
            The setting value.
        """
        value = self.config.get(key, "")
        return value if value else default

    def get_float(self, key: str, default: float) -> float:
        """Read a Config value as a number.

        Args:
            key: Setting name.
            default: Returned when missing or unparseable.

        Returns:
            The numeric value.
        """
        raw = self.config.get(key, "")
        try:
            return float(raw)
        except (TypeError, ValueError):
            return default

    def get_int(self, key: str, default: int) -> int:
        """Read a Config value as a whole number.

        Args:
            key: Setting name.
            default: Returned when missing or unparseable.

        Returns:
            The integer value.
        """
        raw = self.config.get(key, "")
        try:
            return int(float(raw))
        except (TypeError, ValueError):
            return default

    def get_bool(self, key: str, default: bool) -> bool:
        """Read a Config value as a yes/no flag.

        Args:
            key: Setting name.
            default: Returned when missing or unrecognised.

        Returns:
            True for yes/y/true/1, False for no/n/false/0.
        """
        raw = self.config.get(key, "").strip().lower()
        if raw in {"y", "yes", "true", "1"}:
            return True
        if raw in {"n", "no", "false", "0"}:
            return False
        return default

    def get_list(self, key: str, separators: str = ",\n;") -> list[str]:
        """Read a Config value as a list of trimmed, non-empty items.

        Args:
            key: Setting name.
            separators: Characters that separate items.

        Returns:
            The parsed items.
        """
        raw = self.config.get(key, "")
        if not raw:
            return []
        for separator in separators[1:]:
            raw = raw.replace(separator, separators[0])
        return [part.strip() for part in raw.split(separators[0]) if part.strip()]

    def set_config(self, key: str, value: str) -> None:
        """Write a Config value, adding the key if it is not already present.

        Args:
            key: Setting name.
            value: New value, stored as text.
        """
        sheet = self.workbook[CONFIG_SHEET]
        for row in range(2, sheet.max_row + 1):
            if str(sheet.cell(row=row, column=1).value or "").strip() == key:
                sheet.cell(row=row, column=2, value=value)
                return
        sheet.cell(row=sheet.max_row + 1, column=1, value=key)
        sheet.cell(row=sheet.max_row, column=2, value=value)

    @property
    def keywords(self) -> list[Keyword]:
        """Return the configured error-start markers.

        Returns:
            Every marker in the Keywords sheet, with its Enabled flag.
        """
        sheet = self.workbook[KEYWORDS_SHEET]
        markers: list[Keyword] = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            marker = str(row[0]).strip()
            if not marker:
                continue
            flag = str(row[1]).strip().lower() if len(row) > 1 and row[1] is not None else "y"
            markers.append(Keyword(marker, flag in {"y", "yes", "true", "1"}))
        return markers

    @property
    def last_run(self) -> dt.datetime | None:
        """Return the timestamp of the last successful run.

        Returns:
            The stored moment, or None if the script has never completed.
        """
        raw = self.config.get("last_run", "")
        if not raw:
            return None
        try:
            return to_naive_local(raw)
        except TypeError:
            return None

    # -- ledger access -----------------------------------------------------

    def _load_errors(self) -> None:
        """Read the Errors sheet into memory, keyed by signature."""
        sheet = self.workbook[ERRORS_SHEET]
        index = _header_index(sheet, ERROR_HEADERS)
        for row in range(2, sheet.max_row + 1):
            signature = sheet.cell(row=row, column=index["Signature"]).value
            if signature is None or not str(signature).strip():
                continue
            signature = str(signature).strip()
            self._errors[signature] = ErrorRow(
                signature=signature,
                sample=str(sheet.cell(row=row, column=index["Sample Message"]).value or ""),
                total_count=_as_int(sheet.cell(row=row, column=index["Total Count"]).value),
                first_found=_as_datetime(sheet.cell(row=row, column=index["First Found"]).value),
                last_seen=_as_datetime(sheet.cell(row=row, column=index["Last Seen"]).value),
                status=str(sheet.cell(row=row, column=index["Status"]).value or ""),
                case_ids=_split_cell(sheet.cell(row=row, column=index["Case IDs Seen"]).value),
                folders=_split_cell(sheet.cell(row=row, column=index["Source Folders"]).value),
                excel_row=row,
            )

    def _load_entry_ids(self) -> None:
        """Read every Entry ID already recorded, for deduplication."""
        sheet = self.workbook[OCCURRENCES_SHEET]
        index = _header_index(sheet, OCCURRENCE_HEADERS)
        column = index["Entry ID"]
        for row in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row, column=column).value
            if value:
                self._seen_entry_ids.add(str(value))

    @property
    def known_signatures(self) -> list[str]:
        """Return every signature in the ledger.

        Returns:
            Signatures, in no particular order.
        """
        return list(self._errors.keys())

    def already_processed(self, entry_id: str) -> bool:
        """Report whether an email has been processed before.

        Args:
            entry_id: The message's Entry ID.

        Returns:
            True when the id is already in the Occurrences sheet.
        """
        return entry_id in self._seen_entry_ids

    def get_error(self, signature: str) -> ErrorRow | None:
        """Return the ledger row for a signature.

        Args:
            signature: The grouping key.

        Returns:
            The row, or None when the signature is unknown.
        """
        return self._errors.get(signature)

    def is_new_this_run(self, signature: str) -> bool:
        """Report whether a signature was first created during this run.

        Args:
            signature: The grouping key.

        Returns:
            True when this run created the row.
        """
        return signature in self._new_signatures

    def count_this_run(self, signature: str) -> int:
        """Return how many times a signature occurred in this run.

        Args:
            signature: The grouping key.

        Returns:
            The occurrence count, 0 if it did not occur.
        """
        return self._run_counts.get(signature, 0)

    def rows_this_run(self) -> list[ErrorRow]:
        """Return the ledger rows that had at least one occurrence this run.

        This is what the digest email reports on - the ledger as a whole is
        the permanent record, but the draft only describes what just happened.

        Returns:
            The matching rows, in no particular order.
        """
        return [
            record for signature, record in self._errors.items()
            if self._run_counts.get(signature, 0) > 0
        ]

    # -- recording ---------------------------------------------------------

    def begin_run(self) -> None:
        """Reset the per-run counters before processing starts."""
        self._run_counts.clear()
        self._new_signatures.clear()
        self._pending_occurrences.clear()

    def record(
        self,
        signature: str,
        sample_text: str,
        received: dt.datetime,
        folder: str,
        subject: str,
        sender: str,
        case_ids: Iterable[str],
        entry_id: str,
        run_id: str,
    ) -> None:
        """Record one occurrence of one error.

        Creates the ledger row when the signature is new, otherwise bumps the
        counts and Last Seen.  First Found is never altered on an existing row.

        Args:
            signature: The normalized grouping key.
            sample_text: Verbatim error text, kept only for a brand-new row.
            received: When the email arrived.
            folder: Folder the email was found in.
            subject: Email subject.
            sender: Email sender.
            case_ids: Identifiers stripped out of this occurrence.
            entry_id: The email's Entry ID.
            run_id: Identifier of this run.
        """
        ids = [str(value) for value in case_ids]
        existing = self._errors.get(signature)

        if existing is None:
            self._errors[signature] = ErrorRow(
                signature=signature,
                sample=sample_text,
                total_count=1,
                first_found=received,
                last_seen=received,
                status="New",
                case_ids=ids[:],
                folders=[folder] if folder else [],
                excel_row=0,
            )
            self._new_signatures.add(signature)
        else:
            existing.total_count += 1
            if existing.last_seen is None or received > existing.last_seen:
                existing.last_seen = received
            if existing.first_found is None or received < existing.first_found:
                existing.first_found = received
            for value in ids:
                if value not in existing.case_ids:
                    existing.case_ids.append(value)
            if folder and folder not in existing.folders:
                existing.folders.append(folder)
            if not existing.sample:
                existing.sample = sample_text

        self._run_counts[signature] = self._run_counts.get(signature, 0) + 1
        self._seen_entry_ids.add(entry_id)
        self._pending_occurrences.append((
            received, folder, subject, sender, signature,
            sample_text, ", ".join(ids), entry_id, run_id,
        ))

    # -- writing -----------------------------------------------------------

    def save(self, backup: bool = True) -> None:
        """Write every change back to disk.

        Args:
            backup: Copy the previous file to <name>.bak before overwriting.

        Raises:
            PermissionError: If the workbook is open in Excel.  The caller
                should report this and must NOT advance last_run.
        """
        self._write_errors()
        self._append_occurrences()
        if backup and self.path.exists():
            shutil.copy2(self.path, self.path.with_suffix(self.path.suffix + ".bak"))
        self.workbook.save(self.path)

    def _write_errors(self) -> None:
        """Flush the in-memory ledger to the Errors sheet."""
        sheet = self.workbook[ERRORS_SHEET]
        index = _header_index(sheet, ERROR_HEADERS)
        now = dt.datetime.now().replace(microsecond=0)

        # Every existing row starts at zero, so a row that did not occur this
        # run reads 0 rather than keeping last run's figure.
        for row in range(2, sheet.max_row + 1):
            if sheet.cell(row=row, column=index["Signature"]).value:
                sheet.cell(row=row, column=index["Count This Run"], value=0)

        for signature, record in self._errors.items():
            row = record.excel_row
            if not row:
                row = sheet.max_row + 1
                while sheet.cell(row=row, column=index["Signature"]).value:
                    row += 1
                record.excel_row = row
                sheet.cell(row=row, column=index["Signature"], value=signature)
                sheet.cell(row=row, column=index["Sample Message"], value=record.sample)
                sheet.cell(row=row, column=index["Status"], value=record.status or "New")
                first_cell = sheet.cell(row=row, column=index["First Found"],
                                        value=record.first_found)
                first_cell.number_format = _DATE_FORMAT
            elif sheet.cell(row=row, column=index["First Found"]).value is None:
                # Only ever filled when it is genuinely blank.
                first_cell = sheet.cell(row=row, column=index["First Found"],
                                        value=record.first_found)
                first_cell.number_format = _DATE_FORMAT

            run_count = self._run_counts.get(signature, 0)
            if run_count == 0 and signature not in self._new_signatures:
                continue  # untouched this run - leave the row exactly as it is

            sheet.cell(row=row, column=index["Total Count"], value=record.total_count)
            sheet.cell(row=row, column=index["Count This Run"], value=run_count)
            last_cell = sheet.cell(row=row, column=index["Last Seen"], value=record.last_seen)
            last_cell.number_format = _DATE_FORMAT
            sheet.cell(row=row, column=index["Case IDs Seen"],
                       value=", ".join(record.case_ids[-_MAX_CASE_IDS:]))
            sheet.cell(row=row, column=index["Source Folders"],
                       value=", ".join(record.folders))
            updated_cell = sheet.cell(row=row, column=index["Last Updated"], value=now)
            updated_cell.number_format = _DATE_FORMAT
            # Status / Owner / Notes are deliberately absent - see module docs.

    def _append_occurrences(self) -> None:
        """Append this run's occurrence rows."""
        sheet = self.workbook[OCCURRENCES_SHEET]
        for values in self._pending_occurrences:
            sheet.append(values)
            sheet.cell(row=sheet.max_row, column=1).number_format = _DATE_FORMAT
        self._pending_occurrences.clear()

    def prune_occurrences(self, retention_days: int) -> int:
        """Delete occurrence rows older than the retention window.

        The Errors ledger is never pruned - only the per-instance history,
        which is what would otherwise grow without limit.

        Args:
            retention_days: Keep rows received within this many days.  0 or
                less disables pruning.

        Returns:
            The number of rows removed.
        """
        if retention_days <= 0:
            return 0

        sheet = self.workbook[OCCURRENCES_SHEET]
        cutoff = dt.datetime.now() - dt.timedelta(days=retention_days)
        removed = 0
        for row in range(sheet.max_row, 1, -1):
            received = _as_datetime(sheet.cell(row=row, column=1).value)
            if received is not None and received < cutoff:
                sheet.delete_rows(row, 1)
                removed += 1
        return removed


# --------------------------------------------------------------------------
# Sheet construction
# --------------------------------------------------------------------------

def _build_new_workbook() -> Workbook:
    """Create a workbook with all four sheets and the default settings.

    Returns:
        A new openpyxl Workbook, not yet saved.
    """
    workbook = Workbook()
    workbook.remove(workbook.active)

    config = workbook.create_sheet(CONFIG_SHEET)
    _write_header(config, ("Setting", "Value", "What it does"))
    for key, value, description in DEFAULT_CONFIG:
        config.append((key, value, description))
    config.column_dimensions["A"].width = 26
    config.column_dimensions["B"].width = 46
    config.column_dimensions["C"].width = 86
    for row in range(2, config.max_row + 1):
        config.cell(row=row, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        config.cell(row=row, column=3).alignment = Alignment(wrap_text=True, vertical="top")

    keywords = workbook.create_sheet(KEYWORDS_SHEET)
    _write_header(keywords, ("Marker", "Enabled", "Notes"))
    for marker, enabled, note in DEFAULT_KEYWORDS:
        keywords.append((marker, enabled, note))
    keywords.column_dimensions["A"].width = 22
    keywords.column_dimensions["B"].width = 10
    keywords.column_dimensions["C"].width = 60

    errors = workbook.create_sheet(ERRORS_SHEET)
    _write_header(errors, ERROR_HEADERS, user_columns=USER_OWNED_COLUMNS)
    _set_widths(errors, (58, 68, 12, 14, 18, 18, 14, 16, 42, 30, 26, 18))
    errors.freeze_panes = "A2"

    occurrences = workbook.create_sheet(OCCURRENCES_SHEET)
    _write_header(occurrences, OCCURRENCE_HEADERS)
    _set_widths(occurrences, (18, 26, 52, 28, 58, 68, 26, 44, 20))
    occurrences.freeze_panes = "A2"

    return workbook


def _ensure_sheets(workbook: Workbook) -> None:
    """Add any sheet or column missing from an older workbook.

    Lets the workbook gain new columns in a later version without the person
    having to rebuild it and lose their Status and Notes.

    Args:
        workbook: The loaded workbook to repair in place.
    """
    if CONFIG_SHEET not in workbook.sheetnames:
        sheet = workbook.create_sheet(CONFIG_SHEET, 0)
        _write_header(sheet, ("Setting", "Value", "What it does"))
        for key, value, description in DEFAULT_CONFIG:
            sheet.append((key, value, description))
    else:
        sheet = workbook[CONFIG_SHEET]
        present = {
            str(sheet.cell(row=row, column=1).value or "").strip()
            for row in range(2, sheet.max_row + 1)
        }
        for key, value, description in DEFAULT_CONFIG:
            if key not in present:
                sheet.append((key, value, description))

    if KEYWORDS_SHEET not in workbook.sheetnames:
        sheet = workbook.create_sheet(KEYWORDS_SHEET)
        _write_header(sheet, ("Marker", "Enabled", "Notes"))
        for marker, enabled, note in DEFAULT_KEYWORDS:
            sheet.append((marker, enabled, note))

    if ERRORS_SHEET not in workbook.sheetnames:
        sheet = workbook.create_sheet(ERRORS_SHEET)
        _write_header(sheet, ERROR_HEADERS, user_columns=USER_OWNED_COLUMNS)
    else:
        _add_missing_columns(workbook[ERRORS_SHEET], ERROR_HEADERS)

    if OCCURRENCES_SHEET not in workbook.sheetnames:
        sheet = workbook.create_sheet(OCCURRENCES_SHEET)
        _write_header(sheet, OCCURRENCE_HEADERS)
    else:
        _add_missing_columns(workbook[OCCURRENCES_SHEET], OCCURRENCE_HEADERS)


def _add_missing_columns(sheet: Worksheet, headers: Iterable[str]) -> None:
    """Append any header this sheet does not already carry.

    Args:
        sheet: The worksheet to extend.
        headers: The headers the current version expects.
    """
    present = [
        str(sheet.cell(row=1, column=column).value or "").strip()
        for column in range(1, sheet.max_column + 1)
    ]
    for header in headers:
        if header not in present:
            column = len(present) + 1
            cell = sheet.cell(row=1, column=column, value=header)
            cell.fill = _HEADER_FILL
            cell.font = _HEADER_FONT
            present.append(header)


def _write_header(
    sheet: Worksheet,
    headers: Iterable[str],
    user_columns: frozenset[str] = frozenset(),
) -> None:
    """Write and style a header row.

    Args:
        sheet: Target worksheet.
        headers: Header labels, left to right.
        user_columns: Headers to tint, marking them as the person's to edit.
    """
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=column, value=header)
        cell.fill = _USER_FILL if header in user_columns else _HEADER_FILL
        cell.font = Font(bold=True) if header in user_columns else _HEADER_FONT
        cell.alignment = Alignment(vertical="center")


def _set_widths(sheet: Worksheet, widths: Iterable[int]) -> None:
    """Set column widths left to right.

    Args:
        sheet: Target worksheet.
        widths: Width per column.
    """
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width


def _header_index(sheet: Worksheet, expected: Iterable[str]) -> dict[str, int]:
    """Map header names to 1-based column numbers.

    Reads the sheet's actual header row rather than assuming an order, so
    someone reordering columns in Excel does not corrupt the ledger.

    Args:
        sheet: The worksheet to inspect.
        expected: Headers that must exist.

    Returns:
        A dict of header name to column number.

    Raises:
        ValueError: If an expected header is missing from the sheet.
    """
    index: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        name = sheet.cell(row=1, column=column).value
        if name is not None:
            index[str(name).strip()] = column

    missing = [header for header in expected if header not in index]
    if missing:
        raise ValueError(
            f"Sheet '{sheet.title}' is missing column(s): {', '.join(missing)}. "
            "Delete the workbook to regenerate it, or add the column by hand."
        )
    return index


def _as_int(value: object) -> int:
    """Read a cell as a whole number, treating anything unreadable as 0.

    Args:
        value: The raw cell value.

    Returns:
        The integer value.
    """
    try:
        return int(float(value))  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return 0


def _as_datetime(value: object) -> dt.datetime | None:
    """Read a cell as a datetime, tolerating text a person may have typed.

    Args:
        value: The raw cell value.

    Returns:
        The datetime, or None when the cell is blank or unparseable.
    """
    if value is None or value == "":
        return None
    try:
        return to_naive_local(value)
    except TypeError:
        return None


def _split_cell(value: object) -> list[str]:
    """Split a comma-separated cell into a list.

    Args:
        value: The raw cell value.

    Returns:
        The trimmed, non-empty items.
    """
    if not value:
        return []
    return [part.strip() for part in str(value).split(",") if part.strip()]
