"""
The second pass: merge ledger rows that describe the same error.

OPERATOR 2026-11-11: "after the whole email is composed, the email should again
be filtered again should go through regular expression fuzzy search and see
commonality as now i see for many i am getting single occurance in each row for
the same type."

Reproduced before anything was built. Four real-shaped lock-issue messages
whose variable parts are WORDS rather than numbers — DC_CORP_UPDATES,
BATCH_RECON_JOB, SvcOpsDaily, and the WORK-/WRT-/CASE- prefixes, none of which
normalisation touches — scored 0.838-0.874 against each other and produced four
separate single-occurrence rows at the 0.90 live threshold.

The threshold was then chosen by measurement rather than taste: across kinds
the same corpus scored 0.138-0.446, so the gap between 0.446 and 0.838 is wide
and 0.80 sits inside it with room either side.
"""

import datetime as dt

import pytest
from openpyxl import load_workbook

import error_digest
from workbook import DigestWorkbook

LOCK = ("Lock Issue: {ref} is currently being processed by {op}, "
        "please try again later.")
VARIANTS = [("WORK-1859190", "DC_CORP_UPDATES"), ("WRT-1912044", "BATCH_RECON_JOB"),
            ("CASE-2001377", "SvcOpsDaily"), ("WORK-2244981", "NightlyPosting")]


@pytest.fixture()
def book(tmp_path):
    """A workbook holding four split rows for ONE kind of error."""
    path = tmp_path / "d.xlsx"
    error_digest.main(["--workbook", str(path), "--init"])
    digest = DigestWorkbook.open(path)
    digest.begin_run()
    for index, (ref, operator) in enumerate(VARIANTS):
        digest.record(
            signature=f"lock issue: {ref.lower()} is currently being processed "
                      f"by {operator.lower()}, please try again later",
            sample_text=LOCK.format(ref=ref, op=operator),
            received=dt.datetime(2026, 11, 1 + index, 8, 0),
            folder="Inbox\\Alerts", subject="Pega alert", sender="pega@x.com",
            case_ids=[ref], entry_id=f"E{index}", run_id="r",
        )
    digest.save(backup=False)
    return path


def names(path):
    return DigestWorkbook.open(path).known_signatures


def test_the_split_rows_are_the_problem_being_fixed(book):
    """Guards the premise: without the pass there really are four rows."""
    assert len(names(book)) == 4


def test_they_collapse_to_one(book):
    digest = DigestWorkbook.open(book)
    merges = digest.consolidate(0.80)
    assert len(merges) == 3
    assert len(digest.known_signatures) == 1


def test_the_counts_add_up(book):
    digest = DigestWorkbook.open(book)
    digest.consolidate(0.80)
    only = digest.get_error(digest.known_signatures[0])
    assert only.total_count == 4
    assert len(only.case_ids) == 4, "every reference must survive the merge"


def test_first_found_keeps_the_earliest(book):
    digest = DigestWorkbook.open(book)
    digest.consolidate(0.80)
    only = digest.get_error(digest.known_signatures[0])
    assert only.first_found == dt.datetime(2026, 11, 1, 8, 0)
    assert only.last_seen == dt.datetime(2026, 11, 4, 8, 0)


def test_the_merged_row_is_named_by_what_they_share(book):
    digest = DigestWorkbook.open(book)
    digest.consolidate(0.80)
    only = digest.get_error(digest.known_signatures[0])
    assert "is currently being processed by" in only.common
    # The parts that differ between occurrences must be gone.
    for noise in ("DC_CORP_UPDATES", "SvcOpsDaily", "WORK-1859190"):
        assert noise not in only.common


def test_unrelated_errors_are_not_merged(book):
    """The whole risk of a lower threshold, pinned."""
    digest = DigestWorkbook.open(book)
    digest.begin_run()
    digest.record(
        signature="error occured during document generation <n> <n>",
        sample_text="Error occured during document generation 4534 3434",
        received=dt.datetime(2026, 11, 5, 8, 0), folder="Inbox\\Alerts",
        subject="s", sender="x", case_ids=[], entry_id="D1", run_id="r")
    digest.consolidate(0.80)
    assert len(digest.known_signatures) == 2, digest.known_signatures


def test_a_chain_collapses_however_it_arrived(book):
    """
    A~B and B~C but A!~C. One pass over arriving mail can leave A and C apart
    for ever; repeating until a sweep changes nothing cannot.
    """
    digest = DigestWorkbook.open(book)
    merged = digest.consolidate(0.80)
    assert len(digest.known_signatures) == 1
    assert len(merged) == 3, "each merge is reported, not just the final state"


def test_a_threshold_of_zero_switches_the_pass_off(book):
    assert error_digest.main(["--workbook", str(book), "--set",
                              "consolidate_threshold=0"]) == 0
    assert error_digest.main(["--workbook", str(book), "--consolidate"]) == 0
    assert len(names(book)) == 4, "nothing may be merged when the pass is off"


def test_the_persons_triage_is_never_lost(book):
    """A merge deletes a ROW; it must not delete what a person wrote on it."""
    sheet = load_workbook(book)["Errors"]
    head = {sheet.cell(row=1, column=c).value: c
            for c in range(1, sheet.max_column + 1)}
    workbook = load_workbook(book)
    rows = workbook["Errors"]
    rows.cell(row=3, column=head["Status"], value="Open")
    rows.cell(row=3, column=head["Owner"], value="Satya")
    rows.cell(row=3, column=head["Notes"], value="raised with platform team")
    rows.cell(row=4, column=head["Notes"], value="second note")
    workbook.save(book)

    assert error_digest.main(["--workbook", str(book), "--consolidate"]) == 0

    after = load_workbook(book)["Errors"]
    head = {after.cell(row=1, column=c).value: c
            for c in range(1, after.max_column + 1)}
    # "New" is the script's own default and must not outrank real triage —
    # the first implementation kept "New" over "Open".
    assert after.cell(row=2, column=head["Status"]).value == "Open"
    assert after.cell(row=2, column=head["Owner"]).value == "Satya"
    notes = after.cell(row=2, column=head["Notes"]).value
    assert "raised with platform team" in notes and "second note" in notes, (
        "two different notes are kept together, never one chosen over the other")


def test_the_absorbed_rows_are_actually_gone_from_the_sheet(book):
    """A merge that left the old rows behind would double every count."""
    error_digest.main(["--workbook", str(book), "--consolidate"])
    sheet = load_workbook(book)["Errors"]
    populated = sum(1 for r in range(2, sheet.max_row + 1)
                    if sheet.cell(row=r, column=1).value)
    assert populated == 1
