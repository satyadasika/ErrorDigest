"""
End-to-end tests: fixtures in, ledger out.

These exercise the same code path the Windows run uses, with the Outlook
adapter swapped for the fixture reader.  Everything except the ~70 lines of
COM calls is covered here, on any operating system.
"""

import datetime as dt
import json
from pathlib import Path

import pytest

import error_digest
from workbook import DigestWorkbook

FIXTURES = Path(__file__).parent / "fixtures"


def run(workbook_path, fixtures=FIXTURES, since="2026-09-01", extra=()):
    """Run the digest over fixtures with no draft, and return the exit code.

    Args:
        workbook_path: Where the workbook lives.
        fixtures: Fixture directory.
        since: Window start.
        extra: Additional command-line arguments.

    Returns:
        The process exit code.
    """
    return error_digest.main([
        "--workbook", str(workbook_path),
        "--source", "fixtures",
        "--fixtures", str(fixtures),
        "--since", since,
        "--no-draft",
        *extra,
    ])


@pytest.fixture()
def book(tmp_path):
    """Return a path for a fresh workbook in a temporary directory."""
    return tmp_path / "digest.xlsx"


# -- the workbook itself ---------------------------------------------------

def test_workbook_is_created_with_all_four_sheets(book):
    """A first run must produce a usable workbook with no manual setup."""
    assert error_digest.main(["--workbook", str(book), "--init"]) == 0
    digest = DigestWorkbook.open(book)
    assert set(digest.workbook.sheetnames) == {
        "Config", "Keywords", "Errors", "Occurrences"
    }


def test_defaults_are_present_and_readable(book):
    """Every setting the code reads must exist with a sane default."""
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    assert digest.get_float("fuzzy_threshold", 0.0) == 0.90
    assert digest.get_bool("strip_reply_chain", False) is True
    assert digest.get_int("occurrence_retention_days", 0) == 180
    assert any(k.marker == "Error Message" and k.enabled for k in digest.keywords)
    assert any(k.marker == "Warning" and not k.enabled for k in digest.keywords)


def test_a_missing_column_is_added_rather_than_losing_the_file(book):
    """An older workbook gains new columns instead of needing a rebuild."""
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    sheet = digest.workbook["Errors"]
    sheet.delete_cols(sheet.max_column, 1)      # simulate an older layout
    digest.workbook.save(book)
    reopened = DigestWorkbook.open(book)         # must repair, not raise
    headers = [c.value for c in reopened.workbook["Errors"][1]]
    assert "Last Updated" in headers


# -- consolidation ---------------------------------------------------------

def test_errors_consolidate_by_signature(book):
    """Nine fixture emails collapse to five distinct errors."""
    assert run(book) == 0
    digest = DigestWorkbook.open(book)
    assert len(digest.known_signatures) == 5


def test_the_repeated_error_carries_its_count(book):
    """The payment timeout appears three times across three emails."""
    run(book)
    digest = DigestWorkbook.open(book)
    payment = next(
        digest.get_error(s) for s in digest.known_signatures if "payment" in s
    )
    assert payment.total_count == 3
    assert len(payment.case_ids) == 3           # 88421, 99110, 10022


def test_first_found_is_the_earliest_email_not_the_run_time(book):
    """'When did we first see this' must come from the email, not today."""
    run(book)
    digest = DigestWorkbook.open(book)
    payment = next(
        digest.get_error(s) for s in digest.known_signatures if "payment" in s
    )
    assert payment.first_found == dt.datetime(2026, 9, 28, 14, 2)


def test_source_folders_are_recorded(book):
    """The ledger shows which folder an error arrives in."""
    run(book)
    digest = DigestWorkbook.open(book)
    payment = next(
        digest.get_error(s) for s in digest.known_signatures if "payment" in s
    )
    assert "Inbox\\Alerts\\Prod" in payment.folders


def test_quoted_history_is_not_counted_again(book):
    """The forwarded fixture quotes an error that must not be re-counted."""
    run(book)
    digest = DigestWorkbook.open(book)
    assert not any("something quoted" in s for s in digest.known_signatures)


def test_an_email_with_no_errors_adds_nothing(book):
    """The newsletter fixture must not create a row."""
    run(book)
    digest = DigestWorkbook.open(book)
    assert not any("newsletter" in s or "healthy" in s for s in digest.known_signatures)


def test_occurrences_record_every_instance(book):
    """Ten occurrences across five error types."""
    run(book)
    digest = DigestWorkbook.open(book)
    assert digest.workbook["Occurrences"].max_row - 1 == 10


# -- the three durability guarantees --------------------------------------

def test_a_second_run_does_not_double_count(book):
    """Re-reading the same emails must change no count."""
    run(book)
    before = DigestWorkbook.open(book).get_error(
        next(s for s in DigestWorkbook.open(book).known_signatures if "payment" in s)
    ).total_count
    run(book)
    after_digest = DigestWorkbook.open(book)
    signature = next(s for s in after_digest.known_signatures if "payment" in s)
    assert after_digest.get_error(signature).total_count == before == 3


def test_user_columns_survive_a_run(book):
    """Status, Owner and Notes are the person's - the script never writes them."""
    run(book)
    digest = DigestWorkbook.open(book)
    sheet = digest.workbook["Errors"]
    sheet.cell(row=2, column=7, value="Open")
    sheet.cell(row=2, column=8, value="Satya")
    sheet.cell(row=2, column=9, value="Vendor ticket raised")
    digest.workbook.save(book)

    run(book)

    sheet = DigestWorkbook.open(book).workbook["Errors"]
    assert sheet.cell(row=2, column=7).value == "Open"
    assert sheet.cell(row=2, column=8).value == "Satya"
    assert sheet.cell(row=2, column=9).value == "Vendor ticket raised"


def test_first_found_is_never_rewritten(book):
    """Even if the same error arrives again, the original date stands."""
    run(book)
    original = DigestWorkbook.open(book).workbook["Errors"].cell(row=2, column=5).value
    run(book)
    assert DigestWorkbook.open(book).workbook["Errors"].cell(row=2, column=5).value == original


def test_count_this_run_resets_for_an_error_that_did_not_recur(book):
    """A stale 'Count This Run' would misreport last week's noise as today's."""
    run(book)
    run(book)                                   # same emails, all deduped
    sheet = DigestWorkbook.open(book).workbook["Errors"]
    assert sheet.cell(row=2, column=4).value == 0


# -- windowing -------------------------------------------------------------

def test_last_run_advances_after_a_successful_save(book):
    """The next run must start where this one finished."""
    run(book)
    assert DigestWorkbook.open(book).last_run is not None


def test_dry_run_changes_nothing(book):
    """--dry-run must not write the ledger or move the window."""
    error_digest.main(["--workbook", str(book), "--init"])
    run(book, extra=("--dry-run",))
    digest = DigestWorkbook.open(book)
    assert digest.known_signatures == []
    assert digest.last_run is None


def test_since_window_excludes_older_email(book):
    """A later --since must leave the earlier fixtures out."""
    run(book, since="2026-09-29T12:00")
    digest = DigestWorkbook.open(book)
    assert not any(
        digest.get_error(s).first_found < dt.datetime(2026, 9, 29, 12, 0)
        for s in digest.known_signatures
    )


def test_a_bad_since_stops_rather_than_guessing(book):
    """A mistyped date must not silently read the wrong window."""
    error_digest.main(["--workbook", str(book), "--init"])
    with pytest.raises(SystemExit):
        run(book, since="last tuesday")


# -- configuration drives behaviour ---------------------------------------

def test_disabling_a_keyword_changes_what_is_found(book):
    """Proves the Keywords sheet is really the control surface."""
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    sheet = digest.workbook["Keywords"]
    for row in range(2, sheet.max_row + 1):
        if str(sheet.cell(row=row, column=1).value).strip() == "FATAL":
            sheet.cell(row=row, column=2, value="N")
    digest.workbook.save(book)

    run(book)
    assert not any("disk quota" in s for s in DigestWorkbook.open(book).known_signatures)


def test_lowering_the_threshold_merges_more(book):
    """The threshold in Config genuinely drives grouping.

    Measured on these fixtures, the closest two of the five signatures score
    0.400 ("disk quota exceeded on <path>" against "unable to write report to
    <path> for"), so 0.35 is the value that forces a merge.  That the default
    0.90 keeps all five apart is the real result: these errors ARE distinct.
    """
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.set_config("fuzzy_threshold", "0.35")
    digest.workbook.save(book)
    run(book)
    assert len(DigestWorkbook.open(book).known_signatures) < 5


def test_max_errors_per_email_is_honoured(book):
    """Setting it to 1 takes only the first error from the nightly digest."""
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.set_config("max_errors_per_email", "1")
    digest.workbook.save(book)
    run(book)
    assert DigestWorkbook.open(book).workbook["Occurrences"].max_row - 1 == 8


def test_a_custom_keyword_added_by_the_person_works(tmp_path, book):
    """Someone adds 'ALERT CODE' in Excel; it must start matching."""
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.workbook["Keywords"].append(("ALERT CODE", "Y", "added by hand"))
    digest.workbook.save(book)

    custom = tmp_path / "custom"
    custom.mkdir()
    (custom / "a.json").write_text(json.dumps({
        "entry_id": "C1", "received": "2026-09-29T10:00:00",
        "body": "ALERT CODE: Pump pressure out of range for case 9",
    }), encoding="utf-8")

    run(book, fixtures=custom)
    assert any("pump pressure" in s for s in DigestWorkbook.open(book).known_signatures)


# -- failure paths ---------------------------------------------------------

def test_a_malformed_fixture_is_named_not_skipped(tmp_path, book):
    """A silent skip would let a test pass for the wrong reason."""
    broken = tmp_path / "broken"
    broken.mkdir()
    (broken / "bad.json").write_text("{not json", encoding="utf-8")
    assert run(book, fixtures=broken) == 1


def test_a_missing_fixture_directory_fails_clearly(book):
    """A wrong path must report, not produce an empty, reassuring run."""
    assert run(book, fixtures=Path("/no/such/place")) == 1


def test_outlook_run_without_folders_configured_fails(book):
    """Scanning nothing must be an error, not a quiet success."""
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.set_config("folders", "")
    digest.workbook.save(book)
    assert error_digest.main([
        "--workbook", str(book), "--since", "2026-09-01", "--no-draft"
    ]) == 1
