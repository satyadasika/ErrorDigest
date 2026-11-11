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


# -- markup, compound labels, and the group's common wording ---------------
# Added 2026-11-11 after the operator's real ledger was topped by markup:
# "</td>" 21 occurrences, "</pyStepLabel>" 20, "to Appear</CaseStatus>" 17.

def test_closing_tags_are_not_errors():
    """The exact rows that topped the real ledger must not be recordable."""
    for junk in ("</td>", "</pyStepLabel>", "</pyLastFlowStepLabel>", "&nbsp;", "<br/>"):
        assert parsing.is_markup_only(junk), junk


def test_a_real_message_is_not_markup():
    assert not parsing.is_markup_only("Lock Issue: WORK-1 is currently being processed")


def test_a_two_column_table_puts_label_and_value_on_separate_lines():
    html = ('<table><tr><td>Exception/Error Message</td>'
            '<td>No Active Case Case Number 2</td></tr></table>')
    assert parsing.strip_markup(html) == "Exception/Error Message\nNo Active Case Case Number 2"


def test_the_value_cell_of_a_table_is_read():
    html = ('<table><tr><td>Exception/Error Message</td>'
            '<td>No Active Case Case Number 2</td></tr></table>')
    found = parsing.find_error_texts(parsing.strip_markup(html), KEYWORDS)
    assert found == [("Error Message", "No Active Case Case Number 2")]


def test_a_compound_label_does_not_leak_into_the_message():
    """'Exception/Error' left '/Error' behind and it was stored as an error."""
    text = parsing.strip_markup("<table><tr><td>Exception/Error</td><td>Boom happened</td></tr></table>")
    found = parsing.find_error_texts(text, KEYWORDS)
    assert found and found[0][1] == "Boom happened"
    assert not any(t.strip("/ ").lower() in {"error", "exception"} for _m, t in found)


def test_the_value_keeps_its_first_word_even_when_that_word_is_a_marker():
    """'Error occured during...' must not become 'occured during...'."""
    text = parsing.strip_markup(
        "<table><tr><td>Exception/Error</td>"
        "<td>Error occured during document generation 4534 3434</td></tr></table>")
    found = parsing.find_error_texts(text, KEYWORDS)
    assert found[0][1].startswith("Error occured during document generation")
    assert len(found) == 1, f"the value line was scanned twice: {found}"


def test_embedded_xml_keeps_its_words_and_loses_its_tags():
    out = parsing.strip_markup("No Active Case <something>asasas</something> sequence 9")
    assert "asasas" in out and "<" not in out


def test_long_messages_still_group__autojunk():
    """
    difflib's autojunk ignores every character occurring in >1% of a sequence
    longer than 200 chars. On an XML payload that is nearly every character, so
    two occurrences of ONE error measured 0.775 against a true 0.958 and each
    opened its own ledger row. Short messages were unaffected, which is why it
    went unseen.
    """
    pad = ("<Request><Header><System>CRM</System><Version>4.2</Version>"
           "<Timestamp>2026-11-11T09:00:00</Timestamp><Correlation>abc-def-ghi</Correlation>"
           "</Header><Body><Account>{a}</Account><Branch>{b}</Branch>"
           "<Operator>SVC_USER</Operator><Channel>BATCH</Channel>"
           "<Originator>UPSTREAM</Originator></Body></Request>")
    one, _ = signature_of("No Active Case Case Number 2 " + pad.format(a="88421", b="LONDON"))
    two, _ = signature_of("No Active Case Case Number 7 " + pad.format(a="99110", b="PARIS"))
    assert len(one) > 200, "the bug only appears past difflib's 200-character threshold"
    assert similarity(one, two) >= 0.90
    assert best_match(two, [one], 0.90)[0] == one


def test_common_wording_gives_the_operators_names():
    """Their own four examples, verbatim."""
    cases = [
        (["No Active Case Case Number 2 sequence 32",
          "No Active Case Case 3 something else",
          "No Active Case <something>asasas</something> sequence 9"], "No Active Case"),
        (["Process Failed for Case 123 something something as requestor lock",
          "Process Failed for Case 45435 sfdfomething fdfsomething as requestor lock"],
         "Process Failed for Case as requestor lock"),
        (["Error occured during document generation 45345345345 343453455",
          "Error occured during document generation 5546745 3453435345 fgdfg"],
         "Error occured during document generation"),
    ]
    for messages, expected in cases:
        name = messages[0]
        for message in messages[1:]:
            name = parsing.common_wording(name, message)
        assert name == expected, f"{messages[0][:30]}... -> {name!r}"


def test_common_wording_of_one_message_is_that_message():
    assert parsing.common_wording("", "Only one so far") == "Only one so far"


def test_common_wording_can_erode_to_nothing_and_says_so():
    """Two unrelated messages share nothing; the caller must handle an empty name."""
    assert parsing.common_wording("alpha beta", "gamma delta") == ""
