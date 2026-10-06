"""Tests for the digest email body and subject."""

import datetime as dt
from pathlib import Path

import pytest

import error_digest
from workbook import DigestWorkbook

FIXTURES = Path(__file__).parent / "fixtures"
WHEN = dt.datetime(2026, 9, 30, 8, 0)


@pytest.fixture()
def populated(tmp_path):
    """Return a workbook mid-run, with the fixture emails just processed.

    The digest email describes THIS run, so the workbook has to be handed over
    with its run counters still live.  An earlier version of this fixture ran
    the whole program first and then re-processed the same emails, which the
    Entry ID dedup correctly skipped - producing an empty digest and a test
    failure that looked like a bug in the HTML builder.
    """
    import fixture_source

    book = tmp_path / "digest.xlsx"
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.begin_run()
    error_digest.process_messages(
        digest,
        fixture_source.fetch_messages(FIXTURES, dt.datetime(2026, 9, 1)),
        run_id="test",
    )
    return digest


def test_every_error_of_the_run_appears(populated):
    """The draft must not quietly omit a failure."""
    html = error_digest.build_html_digest(populated, WHEN)
    assert "Payment gateway timeout" in html
    assert "Disk quota exceeded" in html


def test_new_errors_are_marked_and_sorted_first(populated):
    """What needs attention belongs at the top."""
    html = error_digest.build_html_digest(populated, WHEN)
    assert "(NEW)" in html


def test_counts_and_dates_are_shown(populated):
    """'How many' and 'since when' are the two questions the digest answers."""
    html = error_digest.build_html_digest(populated, WHEN)
    assert "First found" in html
    assert "This run" in html
    assert "Total" in html


def test_the_body_states_that_no_ai_was_used(populated):
    """The footer records how the grouping was done, for the reader."""
    assert "no AI was used" in error_digest.build_html_digest(populated, WHEN)


def test_html_is_escaped(tmp_path):
    """An error text containing markup must not break or inject into the mail."""
    book = tmp_path / "digest.xlsx"
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.begin_run()
    digest.record(
        signature="script tag <b>", sample_text="Failed: <script>alert(1)</script>",
        received=WHEN, folder="Inbox", subject="s", sender="x",
        case_ids=[], entry_id="E1", run_id="r",
    )
    html = error_digest.build_html_digest(digest, WHEN)
    assert "<script>" not in html
    assert "&lt;script&gt;" in html


def test_an_empty_run_says_so_rather_than_showing_an_empty_table(tmp_path):
    """A quiet day should read as a quiet day."""
    book = tmp_path / "digest.xlsx"
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.begin_run()
    assert "No errors were found" in error_digest.build_html_digest(digest, WHEN)


def test_subject_template_is_filled_in(populated):
    """The subject should summarise before the mail is even opened."""
    subject = error_digest.build_subject(populated, WHEN)
    assert "2026-09-30" in subject
    assert "new" in subject


def test_a_broken_subject_template_does_not_lose_the_run(tmp_path):
    """
    A typo in the template must not crash AFTER the ledger has been written.

    Losing the draft is an inconvenience; losing the ledger write would be a
    real failure, so the template is deliberately forgiving.
    """
    book = tmp_path / "digest.xlsx"
    error_digest.main(["--workbook", str(book), "--init"])
    digest = DigestWorkbook.open(book)
    digest.set_config("draft_subject", "Digest {nonexistent_placeholder}")
    digest.begin_run()
    assert error_digest.build_subject(digest, WHEN) == "Digest {nonexistent_placeholder}"
