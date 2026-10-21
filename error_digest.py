"""
ErrorDigest - consolidate error emails into one Excel ledger and one draft.

Run it with no arguments and it will:

  1. open (or create) error_digest.xlsx
  2. read every email received since the last successful run, from the Outlook
     folders named in the Config sheet
  3. find the configured error markers, strip case/ticket ids, normalize the
     rest, and group identical or near-identical errors
  4. update the Errors ledger - counts, Last Seen, Case IDs - without ever
     touching the Status, Owner or Notes you typed
  5. create an UNSENT draft in Outlook summarising the run and open it

There is no AI anywhere in this program.  Grouping is regular expressions plus
difflib, both from the Python standard library.

Examples:
    python error_digest.py                    normal run, creates the draft
    python error_digest.py --no-draft         scheduled run, ledger only
    python error_digest.py --draft-only       re-create the draft, read nothing
    python error_digest.py --explain          show near-miss match ratios
    python error_digest.py --since 2026-09-01 override the start of the window
    python error_digest.py --dry-run          report, change nothing
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import sys
import uuid
from pathlib import Path

import fixture_source
import outlook_win
from message import EmailMessage
from parsing import (
    DEFAULT_ID_LABELS,
    best_match,
    find_error_texts,
    signature_of,
    strip_reply_chain,
)
from workbook import DigestWorkbook

DEFAULT_WORKBOOK = Path(__file__).resolve().parent / "error_digest.xlsx"
EXIT_OK = 0
EXIT_FAILED = 1


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    """Define every command-line option.

    Returns:
        The configured argument parser.
    """
    parser = argparse.ArgumentParser(
        prog="error_digest",
        description="Consolidate Outlook error emails into an Excel ledger and a draft.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--workbook", type=Path, default=DEFAULT_WORKBOOK,
        help="Path to the Excel workbook (default: error_digest.xlsx beside this script).",
    )
    parser.add_argument(
        "--source", choices=("outlook", "fixtures"), default="outlook",
        help="Where emails come from. 'fixtures' reads JSON files instead of Outlook.",
    )
    parser.add_argument(
        "--fixtures", type=Path,
        help="Directory of fixture emails, required with --source fixtures.",
    )
    parser.add_argument(
        "--since", type=str,
        help="Override the window start, as YYYY-MM-DD or YYYY-MM-DDTHH:MM.",
    )
    parser.add_argument(
        "--no-draft", action="store_true",
        help="Update the ledger but do not create a draft. Use for scheduled runs.",
    )
    parser.add_argument(
        "--draft-only", action="store_true",
        help="Create a draft from what the ledger already holds. Reads no email.",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Report what would happen. The workbook is not written.",
    )
    parser.add_argument(
        "--explain", action="store_true",
        help="Print the similarity ratio for every error, including near misses.",
    )
    parser.add_argument(
        "--init", action="store_true",
        help="Create the workbook with default settings, then stop.",
    )
    parser.add_argument(
        "--list-folders", action="store_true",
        help="Print every Outlook folder path, then stop. Copy one into the "
             "Config sheet's 'folders' setting.",
    )
    return parser


# --------------------------------------------------------------------------
# Pipeline
# --------------------------------------------------------------------------

def process_messages(
    digest: DigestWorkbook,
    messages: list[EmailMessage],
    run_id: str,
    explain: bool = False,
) -> dict[str, int]:
    """Parse every message and record its errors in the ledger.

    Args:
        digest: The open workbook.
        messages: Messages to process, any order.
        run_id: Identifier recorded against every occurrence of this run.
        explain: Print each match and its similarity ratio.

    Returns:
        Counts: emails_read, emails_skipped, occurrences, new_errors.
    """
    keywords = digest.keywords
    threshold = digest.get_float("fuzzy_threshold", 0.90)
    strip_replies = digest.get_bool("strip_reply_chain", True)
    max_per_email = digest.get_int("max_errors_per_email", 0)
    id_labels = digest.get_list("id_labels")

    stats = {"emails_read": 0, "emails_skipped": 0, "occurrences": 0, "new_errors": 0}

    for message in sorted(messages, key=lambda item: item.received):
        if digest.already_processed(message.entry_id):
            stats["emails_skipped"] += 1
            continue

        stats["emails_read"] += 1
        body = strip_reply_chain(message.body) if strip_replies else message.body

        for _marker, error_text in find_error_texts(body, keywords, max_per_email):
            signature, case_ids = signature_of(error_text, id_labels or DEFAULT_ID_LABELS)
            if not signature:
                continue

            matched, ratio = best_match(signature, digest.known_signatures, threshold)
            target = matched or signature

            if explain:
                verdict = "grouped" if matched else "NEW"
                print(f"  [{verdict:7}] {ratio:.3f}  {signature[:70]}")

            was_known = digest.get_error(target) is not None
            digest.record(
                signature=target,
                sample_text=error_text,
                received=message.received,
                folder=message.folder_path,
                subject=message.subject,
                sender=message.sender,
                case_ids=case_ids,
                entry_id=message.entry_id,
                run_id=run_id,
            )
            stats["occurrences"] += 1
            if not was_known:
                stats["new_errors"] += 1

    return stats


def window_start(digest: DigestWorkbook, override: str | None) -> dt.datetime:
    """Decide the moment the scan window begins.

    Precedence: an explicit --since, then the stored last_run, then today minus
    initial_lookback_days for a first ever run.

    Args:
        digest: The open workbook.
        override: The raw --since value, or None.

    Returns:
        The window start as a naive local datetime.

    Raises:
        SystemExit: If --since cannot be parsed, rather than silently reading
            the wrong window.
    """
    if override:
        try:
            return dt.datetime.fromisoformat(override)
        except ValueError:
            raise SystemExit(
                f"--since '{override}' is not a date. Use YYYY-MM-DD or YYYY-MM-DDTHH:MM."
            )

    stored = digest.last_run
    if stored is not None:
        return stored

    days = digest.get_int("initial_lookback_days", 7)
    return dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) \
        - dt.timedelta(days=days)


# --------------------------------------------------------------------------
# The draft email
# --------------------------------------------------------------------------

def build_html_digest(digest: DigestWorkbook, generated_at: dt.datetime) -> str:
    """Render the run's errors as an HTML table.

    New errors are listed first, then recurring ones by descending count, so
    the thing that needs attention is at the top.

    Args:
        digest: The workbook, after processing.
        generated_at: The moment to print in the footer.

    Returns:
        A complete HTML document.
    """
    rows = digest.rows_this_run()
    rows.sort(
        key=lambda record: (
            0 if digest.is_new_this_run(record.signature) else 1,
            -digest.count_this_run(record.signature),
        )
    )

    if not rows:
        body = (
            "<p style='font:14px Segoe UI,Arial;color:#444'>"
            "No errors were found in this run.</p>"
        )
    else:
        cells = []
        for record in rows:
            is_new = digest.is_new_this_run(record.signature)
            first_found = _format_date(record.first_found)
            if is_new:
                first_found = f"<b>{first_found} (NEW)</b>"
            cells.append(
                "<tr>"
                f"<td style='{_TD}'>{html.escape(_display_text(record))}</td>"
                f"<td style='{_TD_NUM}'>{digest.count_this_run(record.signature)}</td>"
                f"<td style='{_TD_NUM}'>{record.total_count}</td>"
                f"<td style='{_TD}'>{first_found}</td>"
                f"<td style='{_TD}'>{_format_date(record.last_seen)}</td>"
                f"<td style='{_TD}'>{html.escape(record.status or '')}</td>"
                f"<td style='{_TD_SMALL}'>{html.escape(', '.join(record.case_ids[-8:]))}</td>"
                "</tr>"
            )
        body = (
            "<table style='border-collapse:collapse;font:13px Segoe UI,Arial'>"
            "<tr>"
            f"<th style='{_TH}'>Error (example text)</th>"
            f"<th style='{_TH}'>This run</th>"
            f"<th style='{_TH}'>Total</th>"
            f"<th style='{_TH}'>First found</th>"
            f"<th style='{_TH}'>Last seen</th>"
            f"<th style='{_TH}'>Status</th>"
            f"<th style='{_TH}'>Recent IDs</th>"
            "</tr>" + "".join(cells) + "</table>"
        )

    new_count = sum(1 for record in rows if digest.is_new_this_run(record.signature))
    recurring = len(rows) - new_count
    occurrences = sum(digest.count_this_run(record.signature) for record in rows)

    return (
        "<html><body style='font:14px Segoe UI,Arial;color:#222'>"
        f"<p>{new_count} new error type(s), {recurring} recurring, "
        f"{occurrences} occurrence(s) in total.</p>"
        f"{body}"
        "<p style='font:12px Segoe UI,Arial;color:#777;margin-top:16px'>"
        f"Generated {generated_at:%Y-%m-%d %H:%M} from "
        f"{html.escape(digest.path.name)}. Grouping is rule-based "
        f"(threshold {digest.get_float('fuzzy_threshold', 0.90):.2f}); no AI was used."
        "</p></body></html>"
    )


_TH = ("border:1px solid #ccc;padding:6px 10px;background:#1F3864;"
       "color:#fff;text-align:left")
_TD = "border:1px solid #ccc;padding:6px 10px;vertical-align:top"
_TD_NUM = "border:1px solid #ccc;padding:6px 10px;text-align:right"
_TD_SMALL = "border:1px solid #ccc;padding:6px 10px;font-size:11px;color:#555"


def _display_text(record) -> str:
    """Choose the clearest text to show for an error.

    The verbatim sample reads better than the normalized signature, so it is
    preferred; the signature is the fallback.

    Args:
        record: An ErrorRow.

    Returns:
        Text for the Error column, trimmed to a readable length.
    """
    text = record.sample or record.signature
    return text if len(text) <= 160 else text[:157] + "..."


def _format_date(value: dt.datetime | None) -> str:
    """Format a datetime for the digest table.

    Args:
        value: The moment, or None.

    Returns:
        "YYYY-MM-DD HH:MM", or "-" when there is nothing to show.
    """
    return f"{value:%Y-%m-%d %H:%M}" if value else "-"


def build_subject(digest: DigestWorkbook, generated_at: dt.datetime) -> str:
    """Fill in the subject template from Config.

    Args:
        digest: The workbook, after processing.
        generated_at: Used for {date}.

    Returns:
        The subject line.  An unknown placeholder leaves the template as-is
        rather than crashing the run after the ledger has been written.
    """
    signatures = [s for s in digest.known_signatures if digest.count_this_run(s) > 0]
    new_count = sum(1 for s in signatures if digest.is_new_this_run(s))
    template = digest.get_str(
        "draft_subject", "Error Digest {date} - {new} new, {recurring} recurring"
    )
    try:
        return template.format(
            date=f"{generated_at:%Y-%m-%d}",
            new=new_count,
            recurring=len(signatures) - new_count,
            total=sum(digest.count_this_run(s) for s in signatures),
        )
    except (KeyError, IndexError, ValueError):
        return template


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    """Run the digest.

    Args:
        argv: Command-line arguments, or None to read sys.argv.

    Returns:
        0 on success, 1 on a handled failure.
    """
    args = build_parser().parse_args(argv)
    started_at = dt.datetime.now().replace(microsecond=0)
    run_id = f"{started_at:%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}"

    try:
        digest = DigestWorkbook.open(args.workbook)
    except (OSError, ValueError) as exc:
        print(f"Could not open {args.workbook}: {exc}", file=sys.stderr)
        return EXIT_FAILED

    if args.list_folders:
        # Printed before anything else touches the workbook, so this works even
        # on a machine where nothing has been configured yet.
        try:
            paths = outlook_win.list_folders()
        except outlook_win.OutlookUnavailable as exc:
            print(f"Could not read your folders: {exc}", file=sys.stderr)
            return EXIT_FAILED
        print("Your Outlook folders — copy one of these into the Config sheet:\n")
        for path in paths:
            print(f"    {path}")
        print(f"\n{len(paths)} folder(s). Several folders go one per line in that "
              "cell (Alt+Enter inside Excel).")
        return EXIT_OK

    if args.init:
        if not args.workbook.exists():
            digest.save(backup=False)
        print(f"Workbook ready: {args.workbook}")
        print("Open it and set 'folders' and 'draft_to' on the Config sheet.")
        return EXIT_OK

    digest.begin_run()
    stats = {"emails_read": 0, "emails_skipped": 0, "occurrences": 0, "new_errors": 0}

    if not args.draft_only:
        since = window_start(digest, args.since)
        print(f"Reading email received since {since:%Y-%m-%d %H:%M}")

        try:
            messages = _collect_messages(digest, args, since)
        except (outlook_win.OutlookUnavailable, FileNotFoundError, ValueError) as exc:
            print(f"Could not read email: {exc}", file=sys.stderr)
            return EXIT_FAILED

        print(f"{len(messages)} message(s) in the window")
        stats = process_messages(digest, messages, run_id, explain=args.explain)
        print(
            f"  processed {stats['emails_read']}, "
            f"already seen {stats['emails_skipped']}, "
            f"errors found {stats['occurrences']}, "
            f"new error types {stats['new_errors']}"
        )

        if args.dry_run:
            print("Dry run - the workbook was not written and last_run was not advanced.")
        else:
            removed = digest.prune_occurrences(
                digest.get_int("occurrence_retention_days", 180)
            )
            if removed:
                print(f"  pruned {removed} occurrence row(s) past the retention window")

            # last_run only advances on a SUCCESSFUL save.  If the save fails
            # the same window is re-read next time, and the Entry ID dedup
            # makes that harmless.
            digest.set_config("last_run", started_at.isoformat(timespec="seconds"))
            try:
                digest.save()
            except PermissionError:
                print(
                    f"\nCould not write {args.workbook} - it is open in Excel.\n"
                    "Close the workbook and run again. Nothing was lost: last_run was "
                    "not advanced, so the same emails will be read next time.",
                    file=sys.stderr,
                )
                return EXIT_FAILED
            except OSError as exc:
                print(f"\nCould not write {args.workbook}: {exc}", file=sys.stderr)
                return EXIT_FAILED
            print(f"  ledger saved to {args.workbook}")

    if args.no_draft:
        return EXIT_OK

    if args.draft_only:
        print("Draft-only: no email was read; the draft reflects the last run's figures.")

    html_body = build_html_digest(digest, started_at)
    subject = build_subject(digest, started_at)

    if args.dry_run:
        print(f"\nDry run - the draft was not created. It would read:\n  {subject}")
        return EXIT_OK

    try:
        outlook_win.create_draft(
            to=digest.get_str("draft_to"),
            cc=digest.get_str("draft_cc"),
            subject=subject,
            html_body=html_body,
            display=True,
        )
    except outlook_win.OutlookUnavailable as exc:
        fallback = args.workbook.with_name("last_digest.html")
        fallback.write_text(html_body, encoding="utf-8")
        print(
            f"\nCould not create the Outlook draft: {exc}\n"
            f"The ledger is saved. The digest was written to {fallback} instead.",
            file=sys.stderr,
        )
        return EXIT_FAILED

    print(f"\nDraft created and opened in Outlook: {subject}")
    print("It is NOT sent. Review it and click Send yourself.")
    return EXIT_OK


def _collect_messages(
    digest: DigestWorkbook,
    args: argparse.Namespace,
    since: dt.datetime,
) -> list[EmailMessage]:
    """Fetch messages from whichever source was chosen.

    Args:
        digest: The open workbook, read for the folder list.
        args: Parsed command-line arguments.
        since: Window start.

    Returns:
        The messages to process.

    Raises:
        SystemExit: If --source fixtures was given without --fixtures.
        ValueError: If no folders are configured for an Outlook run.
    """
    if args.source == "fixtures":
        if not args.fixtures:
            raise SystemExit("--source fixtures needs --fixtures <directory>")
        return fixture_source.fetch_messages(args.fixtures, since)

    folders = digest.get_list("folders", separators="\n,;")
    if not folders:
        raise ValueError(
            "No folders configured. Open the workbook, Config sheet, and set "
            "'folders' to the Outlook folder path(s) to scan, e.g. Inbox\\Alerts"
        )
    return outlook_win.fetch_messages(folders, since)


if __name__ == "__main__":
    sys.exit(main())
