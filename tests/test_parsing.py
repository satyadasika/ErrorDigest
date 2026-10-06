"""Tests for the deterministic extraction and grouping rules."""

import parsing
from parsing import (
    Keyword,
    best_match,
    extract_ids,
    find_error_texts,
    normalize,
    signature_of,
    similarity,
    strip_reply_chain,
)

KEYWORDS = [
    Keyword("Error Message"),
    Keyword("Exception"),
    Keyword("Error"),
    Keyword("FATAL"),
    Keyword("Warning", enabled=False),
]


# -- markers ---------------------------------------------------------------

def test_longest_marker_wins():
    """'Error Message' must win over the shorter 'Error' on the same line."""
    found = find_error_texts("Error Message: disk full", KEYWORDS)
    assert found == [("Error Message", "disk full")]


def test_marker_on_its_own_line_takes_the_next_line():
    """Machine alerts often put the text on the following line."""
    body = "Error Message:\n\nPayment gateway timeout\n"
    assert find_error_texts(body, KEYWORDS) == [("Error Message", "Payment gateway timeout")]


def test_disabled_marker_is_ignored():
    """A marker set to N in the Keywords sheet must never match."""
    assert find_error_texts("Warning: disk is filling up", KEYWORDS) == []


def test_marker_does_not_match_inside_a_word():
    """'Error' must not fire on 'Errors' or 'Terror' - whole words only."""
    assert find_error_texts("Terrors of the night", KEYWORDS) == []
    assert find_error_texts("Errorsome text here", KEYWORDS) == []


def test_every_marker_in_a_body_is_found():
    """A digest email bundling several failures yields several errors."""
    body = "Error Message: one\nFATAL two\nException: three"
    assert len(find_error_texts(body, KEYWORDS)) == 3


def test_max_errors_per_email_caps_the_result():
    """max_errors_per_email=1 keeps only the first match."""
    body = "Error Message: one\nFATAL two"
    assert len(find_error_texts(body, KEYWORDS, max_errors=1)) == 1


def test_case_insensitive_marker():
    """Markers match whatever case the sender used."""
    assert find_error_texts("ERROR MESSAGE: boom", KEYWORDS)[0][1] == "boom"


# -- reply chains ----------------------------------------------------------

def test_quoted_history_is_dropped():
    """Errors below 'Original Message' belong to an email already counted."""
    body = (
        "Exception: new problem\n"
        "-----Original Message-----\n"
        "Error Message: old problem\n"
    )
    kept = strip_reply_chain(body)
    assert "new problem" in kept
    assert "old problem" not in kept


def test_from_header_also_starts_the_quoted_part():
    """Outlook forwards sometimes begin with a bare From: line."""
    body = "Exception: new\nFrom: someone@example.com\nError: old\n"
    assert "old" not in strip_reply_chain(body)


def test_body_without_history_is_untouched():
    """Nothing is lost from an email that quotes nothing."""
    body = "Error Message: standalone problem"
    assert strip_reply_chain(body) == body


# -- identifiers -----------------------------------------------------------

def test_case_id_is_captured_and_removed():
    """The headline behaviour: the case number leaves the signature."""
    cleaned, ids = extract_ids("Payment failed: case ID 12345")
    assert ids == ["12345"]
    assert "12345" not in cleaned


def test_several_id_styles():
    """Labels, hashes, colons and 'no' forms are all identifiers."""
    for text, expected in (
        ("failed for ticket #7781", "7781"),
        ("job no: 42 aborted", "42"),
        ("request id = 4f2c9", "4f2c9"),
        ("incident 2026-993 raised", "2026-993"),
    ):
        _, ids = extract_ids(text)
        assert expected in ids, text


def test_words_that_merely_look_like_labels_are_not_ids():
    """'case sensitive' has no digits, so it is prose, not an identifier."""
    cleaned, ids = extract_ids("Comparison is case sensitive and order matters")
    assert ids == []
    assert cleaned == "Comparison is case sensitive and order matters"


def test_two_case_ids_produce_one_signature():
    """The whole point: same failure, different case numbers, one group."""
    first, _ = signature_of("Payment gateway timeout: case ID 88421")
    second, _ = signature_of("Payment gateway timeout: case ID 99110")
    assert first == second


# -- normalisation ---------------------------------------------------------

def test_variable_parts_become_placeholders():
    """Each kind of variable text maps to its own placeholder."""
    assert "<guid>" in normalize("failed for 3f2504e0-4f89-11d3-9a0c-0305e82c3301")
    assert "<date>" in normalize("run of 2026-09-28 was rejected")
    assert "<path>" in normalize("cannot read C:\\Logs\\app.txt")
    assert "<url>" in normalize("posting to https://api.example.com/v1/pay")
    assert "<email>" in normalize("no mailbox for user@example.com")
    assert "<n>" in normalize("retried 17 times")


def test_a_trailing_timestamp_does_not_split_an_error():
    """Found in the fixtures: a sender appends a time only sometimes."""
    with_time, _ = signature_of("Payment gateway timeout: case ID 1 at 2026-09-28 14:02")
    without, _ = signature_of("Payment gateway timeout: case ID 2")
    assert with_time == without


def test_a_trailing_date_clause_is_removed_entirely():
    """"stopped on <date>" carries no information about WHICH error it is."""
    assert normalize("Service stopped on 2026-09-28") == "service stopped"


def test_a_meaningful_trailing_number_is_kept():
    """'after 30 seconds' is description, not a timestamp - it must survive."""
    signature, _ = signature_of("Connection timeout after 30 seconds")
    assert signature.endswith("seconds")


def test_a_wholly_quoted_message_is_not_collapsed():
    """Replacing the entire text with <q> would group unrelated errors."""
    first = normalize('"Customer record could not be located"')
    second = normalize('"Payment authorisation was declined"')
    assert first != second


def test_a_quoted_value_inside_a_sentence_does_generalise():
    """A quoted value is variable, so it should group."""
    first, _ = signature_of("User 'alice' is not permitted")
    second, _ = signature_of("User 'bob' is not permitted")
    assert first == second


# -- fuzzy grouping --------------------------------------------------------

def test_wording_drift_groups_above_the_threshold():
    """'timeout' vs 'timed out' is the same failure."""
    first, _ = signature_of("Connection timeout on host SRV-12")
    second, _ = signature_of("Connection timed out on host SRV-44")
    assert similarity(first, second) >= 0.90
    assert best_match(second, [first], 0.90)[0] == first


def test_unrelated_errors_do_not_group():
    """Two genuinely different failures must stay apart."""
    first, _ = signature_of("Payment gateway timeout")
    second, _ = signature_of("Disk quota exceeded on the log volume")
    assert best_match(second, [first], 0.90)[0] is None


def test_exact_match_is_preferred_over_a_close_one():
    """An identical signature always wins, whatever else is near."""
    known = ["payment gateway timeout", "payment gateway timeouts"]
    assert best_match("payment gateway timeout", known, 0.90) == ("payment gateway timeout", 1.0)


def test_near_miss_ratio_is_reported_for_explain():
    """--explain needs the ratio even when nothing matched."""
    matched, ratio = best_match("completely different text", ["payment gateway timeout"], 0.90)
    assert matched is None
    assert 0.0 < ratio < 0.90


def test_threshold_is_respected():
    """Lowering the threshold groups more - that is what the setting is for."""
    first, _ = signature_of("Connection refused by broker")
    second, _ = signature_of("Connection rejected by broker")
    assert best_match(second, [first], 0.99)[0] is None
    assert best_match(second, [first], 0.80)[0] == first


def test_empty_input_is_safe():
    """Blank bodies and blank signatures must not raise."""
    assert find_error_texts("", KEYWORDS) == []
    assert strip_reply_chain("") == ""
    assert normalize("") == ""
    assert signature_of("") == ("", [])
    assert best_match("", ["x"], 0.9) == (None, 0.0)


def test_id_labels_are_configurable():
    """The workbook can add a label the default list does not carry."""
    _, ids = extract_ids("failed for workorder 55", labels=["workorder"])
    assert ids == ["55"]
