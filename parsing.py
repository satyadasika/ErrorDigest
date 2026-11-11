"""
Deterministic error extraction and grouping.

There is no AI in this module and there must never be one.  Everything here is
regular expressions and difflib (Python's standard-library string similarity),
so the same input always produces the same output and the result can be
explained line by line.

The pipeline, in order:

    raw body
      -> strip_reply_chain()   drop quoted history so a forwarded alert does
                               not re-count errors that were already counted
      -> find_error_texts()    locate the configured markers ("Error Message:",
                               "Exception:", ...) and take what follows
      -> extract_ids()         pull out case/ticket/request ids and REMOVE them
      -> normalize()           replace dates, numbers, paths, guids, quoted
                               values with placeholders
      -> signature             the stable string used for grouping
      -> best_match()          exact match first, then fuzzy above a threshold
"""

from __future__ import annotations

import html as _html
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Iterable, Sequence


# --------------------------------------------------------------------------
# Keywords
# --------------------------------------------------------------------------

@dataclass(frozen=True)
class Keyword:
    """One configurable error-start marker read from the workbook.

    Attributes:
        marker: The literal text that introduces an error, e.g. "Error Message".
            Matched case-insensitively and only on whole words.
        enabled: Whether this marker participates in matching.
    """

    marker: str
    enabled: bool = True


# Separators allowed between a marker and the error text itself.
# "/" matters: a label cell reading "Exception/Error" left "/Error" behind and
# that was stored as an error in its own right, twice, in the operator's real
# ledger. "," and ";" appear in the same role.
_MARKER_SEPARATOR = re.compile(r"^\s*[:\-–—>|/,;]*\s*")


def _marker_pattern(marker: str) -> re.Pattern[str]:
    """Build a whole-word, case-insensitive pattern for one marker.

    The lookarounds stop "Error" from matching inside "Errors" or "Terror",
    which would otherwise attribute junk lines to a configured marker.

    Args:
        marker: The literal marker text.

    Returns:
        A compiled regular expression.
    """
    return re.compile(
        r"(?<![A-Za-z])" + re.escape(marker.strip()) + r"(?![A-Za-z])",
        re.IGNORECASE,
    )


# --------------------------------------------------------------------------
# Markup
# --------------------------------------------------------------------------

# Tags whose END means a new line of text: table cells and rows, breaks, blocks.
# Converting these to newlines BEFORE dropping the rest is what makes a
# two-column table readable — "<td>Error Message</td><td>No Active Case</td>"
# becomes two lines, and the existing "marker alone on its line -> take the next
# line" rule then picks the value out of the right-hand cell with no new logic.
_LINE_TAGS = re.compile(
    r"</?\s*(?:td|th|tr|table|tbody|thead|br|p|div|li|ul|ol|h[1-6])\b[^>]*>",
    re.IGNORECASE,
)
_ANY_TAG = re.compile(r"<[^>]{0,400}>")
_SCRIPT_OR_STYLE = re.compile(r"<(script|style)\b[^>]*>.*?</\1\s*>", re.IGNORECASE | re.DOTALL)
# A fragment that is only markup, punctuation or entity noise is not an error.
_HAS_LETTERS = re.compile(r"[A-Za-z]")


def strip_markup(body: str) -> str:
    """Turn an HTML/XML-bearing body into plain lines of text.

    WHY THIS EXISTS, from the operator's own ledger on 2026-11-11: their Pega
    alerts carry case data as markup inside the body, so a line scan hit a
    marker somewhere in the soup and took the rest of the line. The top rows of
    the Errors sheet were "</td>" (21 occurrences), "</pyStepLabel>" (20),
    "to Appear</CaseStatus>" (17) and a raw <table border="1" ...> fragment —
    markup counted and ranked as though it were a failure.

    Tags that END a visual line become newlines first, so a table's cells land
    on separate lines; every other tag is removed; entities are decoded. The
    words inside custom XML tags are KEPT (<something>asasas</something> ->
    asasas), because the error phrase often sits inside them and whatever is
    merely payload drops out later when the group's common wording is taken.

    Args:
        body: The raw body, which may be plain text, HTML, XML, or a mixture.

    Returns:
        Plain text with line structure preserved. Input with no markup is
        returned with only its entities decoded and its blank runs collapsed.
    """
    if not body:
        return ""

    text = _SCRIPT_OR_STYLE.sub(" ", body)
    text = _LINE_TAGS.sub("\n", text)
    text = _ANY_TAG.sub(" ", text)
    text = _html.unescape(text)
    # Non-breaking spaces survive unescape as U+00A0 and would defeat \s+.
    text = text.replace("\u00a0", " ").replace("\u200b", "")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


def is_markup_only(text: str) -> bool:
    """Report whether a fragment carries no readable words.

    Used to refuse an extracted "error" that is really a stray tag or a run of
    punctuation. Refusing is the right answer rather than storing it and hoping
    a person notices: a junk row with 21 occurrences outranks real failures in
    the digest.

    Args:
        text: A candidate error message.

    Returns:
        True when there is nothing worth recording.
    """
    if not text or not text.strip():
        return True
    cleaned = _ANY_TAG.sub(" ", text)
    cleaned = _html.unescape(cleaned)
    letters = _HAS_LETTERS.findall(cleaned)
    # Two letters is not a sentence; "td", "br" and "p" are all shorter than the
    # shortest real message seen in the operator's ledger.
    return len(letters) < 3


# --------------------------------------------------------------------------
# Reply-chain stripping
# --------------------------------------------------------------------------

_REPLY_MARKERS: tuple[re.Pattern[str], ...] = (
    re.compile(r"^\s*-{2,}\s*original message\s*-{2,}\s*$", re.IGNORECASE),
    re.compile(r"^\s*_{5,}\s*$"),
    re.compile(r"^\s*-{5,}\s*$"),
    re.compile(r"^\s*from:\s.+$", re.IGNORECASE),
    re.compile(r"^\s*sent:\s.+$", re.IGNORECASE),
    re.compile(r"^\s*on\s.{4,80}\swrote:\s*$", re.IGNORECASE),
    re.compile(r"^\s*>+\s?.*$"),
)


def strip_reply_chain(body: str) -> str:
    """Return only the newest part of an email body.

    Everything from the first quoted-history marker onwards is discarded, so a
    forwarded alert contributes its own errors and not the errors of every
    message it quotes.

    Args:
        body: The full plain-text body.

    Returns:
        The body text up to (not including) the first reply marker.  If no
        marker is present the body is returned unchanged.
    """
    if not body:
        return ""

    lines = body.splitlines()
    for index, line in enumerate(lines):
        for pattern in _REPLY_MARKERS:
            if pattern.match(line):
                return "\n".join(lines[:index])
    return body


# --------------------------------------------------------------------------
# Finding the error text
# --------------------------------------------------------------------------

def _eat_label_run(text: str, compiled: Sequence[tuple[str, "re.Pattern[str]"]]) -> str:
    """Consume any further label words that are part of the same label.

    "Exception/Error Message: X" is ONE label followed by X, not the label
    "Exception" followed by "/Error Message: X". This walks forward while the
    text still begins with a separator and another known marker.

    Args:
        text: What follows the marker that matched, separator already trimmed.
        compiled: The (marker, pattern) pairs, longest first.

    Returns:
        The text with the rest of the label removed.
    """
    for _ in range(4):          # four joined labels is already far-fetched
        progressed = False
        for _marker, pattern in compiled:
            match = pattern.match(text)
            if match and match.start() == 0:
                text = _MARKER_SEPARATOR.sub("", text[match.end():]).strip()
                progressed = True
                break
        if not progressed:
            return text
    return text


def find_error_texts(
    body: str,
    keywords: Sequence[Keyword],
    max_errors: int = 0,
    max_chars: int = 0,
) -> list[tuple[str, str]]:
    """Find every configured error marker in a body and return what follows.

    Markers are tried longest-first, so a line beginning "Error Message: X" is
    attributed to "Error Message" rather than to the shorter "Error".  When a
    marker ends its line the next non-empty line is taken as the error text,
    which is the common layout of machine-generated alerts:

        Error Message:
        Payment gateway timeout

    Args:
        body: Body text, normally already passed through strip_reply_chain().
        keywords: Configured markers; disabled ones are ignored.
        max_errors: Maximum matches to return.  0 means no limit.
        max_chars: Trim each message to this many characters.  0 means no trim.
            A Pega alert can carry kilobytes of embedded case data after the
            sentence that names the failure, and that tail is payload, not
            identity.

    Returns:
        A list of (marker, error_text) pairs in the order they appear.  Empty
        if nothing matched.
    """
    enabled = [k for k in keywords if k.enabled and k.marker.strip()]
    if not enabled or not body:
        return []

    ordered = sorted(enabled, key=lambda k: len(k.marker), reverse=True)
    compiled = [(k.marker, _marker_pattern(k.marker)) for k in ordered]

    lines = body.splitlines()
    found: list[tuple[str, str]] = []
    # Lines already consumed as the VALUE of a label on an earlier line. In a
    # two-column table the label and its value are separate lines, and the value
    # must be taken WHOLE: scanning it again strips its first word whenever that
    # word is itself a marker, turning "Error occured during document
    # generation" into "occured during document generation".
    taken: set[int] = set()

    for index, line in enumerate(lines):
        if index in taken:
            continue
        for marker, pattern in compiled:
            match = pattern.search(line)
            if not match:
                continue

            remainder = line[match.end():]
            remainder = _MARKER_SEPARATOR.sub("", remainder).strip()

            if not remainder:
                value_index = _next_non_empty_index(lines, index + 1)
                if value_index is not None:
                    remainder = lines[value_index].strip()
                    taken.add(value_index)

            # A LABEL CAN BE A RUN OF LABELS. "Exception/Error Message:" is one
            # label, but longest-first matching takes "Exception" and leaves
            # "/Error Message: ..." as the message. Measured on the real string:
            # "Exception/Error: X" came out as "/Error: X". So keep eating any
            # further markers joined by a separator before reading the value.
            remainder = _eat_label_run(remainder, compiled)
            if not remainder:
                value_index = _next_non_empty_index(lines, index + 1)
                if value_index is not None:
                    remainder = lines[value_index].strip()
                    taken.add(value_index)

            if remainder and not is_markup_only(remainder):
                if max_chars and len(remainder) > max_chars:
                    remainder = remainder[:max_chars].rstrip()
                found.append((marker, remainder))
                if max_errors and len(found) >= max_errors:
                    return found
            break  # one marker per line - the longest that matched

    return found


def _next_non_empty_index(lines: Sequence[str], start: int) -> int | None:
    """Return the index of the first non-blank line at or after a position.

    The INDEX rather than the text, so the caller can mark that line consumed
    and not scan it again as a line in its own right.

    Args:
        lines: All body lines.
        start: Index to begin looking from.

    Returns:
        The index, or None if every remaining line is blank.
    """
    for offset, line in enumerate(lines[start:]):
        if line.strip():
            return start + offset
    return None


# --------------------------------------------------------------------------
# Identifier capture - these are REMOVED from the signature
# --------------------------------------------------------------------------

DEFAULT_ID_LABELS: tuple[str, ...] = (
    "case", "ticket", "incident", "request", "correlation", "transaction",
    "job", "order", "reference", "ref", "session", "batch", "trace",
    "record", "invoice", "claim", "message", "event",
)


def _id_pattern(labels: Iterable[str]) -> re.Pattern[str]:
    """Build the pattern that finds labelled identifiers.

    Matches forms such as "case ID 12345", "ticket #A-77", "job no: 9",
    "request id = 4f2c".  The captured value must contain at least one digit,
    which is checked after matching so that phrases like "case sensitive" and
    "order matters" are never mistaken for identifiers.

    Args:
        labels: The identifier label words, from Config.

    Returns:
        A compiled regular expression with named groups 'label' and 'value'.
    """
    alternation = "|".join(re.escape(label.strip()) for label in labels if label.strip())
    return re.compile(
        r"(?<![A-Za-z])"
        r"(?P<label>" + alternation + r")"
        r"\s*(?:id|ids|identifier|number|num|no|#)?"
        r"\s*(?:[:#=]|\bis\b)?\s*"
        r"(?P<value>[A-Za-z0-9][A-Za-z0-9_.\-]*)",
        re.IGNORECASE,
    )


_BARE_HASH_ID = re.compile(r"(?<![A-Za-z0-9])#\s?(?P<value>\d[A-Za-z0-9_\-]*)")
_HAS_DIGIT = re.compile(r"\d")


def extract_ids(
    text: str,
    labels: Iterable[str] = DEFAULT_ID_LABELS,
) -> tuple[str, list[str]]:
    """Remove labelled identifiers from an error text and return them.

    This is the step that makes "Payment failed: case ID 12345" and
    "Payment failed: case ID 98765" the same error.  The identifiers are not
    thrown away - they are returned so the ledger can record which cases an
    error has been seen on.

    Args:
        text: The error text, straight after the marker.
        labels: Identifier label words to look for.

    Returns:
        A tuple of (text with the identifier phrases removed, list of the
        identifier values found, in order, without duplicates).
    """
    if not text:
        return "", []

    found: list[str] = []

    def _take_labelled(match: re.Match[str]) -> str:
        value = match.group("value")
        if not _HAS_DIGIT.search(value):
            return match.group(0)  # "case sensitive" - not an identifier
        if value not in found:
            found.append(value)
        return " "

    cleaned = _id_pattern(labels).sub(_take_labelled, text)

    def _take_hash(match: re.Match[str]) -> str:
        value = match.group("value")
        if value not in found:
            found.append(value)
        return " "

    cleaned = _BARE_HASH_ID.sub(_take_hash, cleaned)
    return _tidy(cleaned), found


# --------------------------------------------------------------------------
# Normalisation
# --------------------------------------------------------------------------

_GUID = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
    r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
)
_ISO_DATETIME = re.compile(
    r"\b\d{4}-\d{2}-\d{2}[T ]\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?\b"
)
_DATE = re.compile(
    r"\b(?:\d{4}-\d{2}-\d{2}|\d{1,2}[/.]\d{1,2}[/.]\d{2,4})\b"
)
_TIME = re.compile(r"\b\d{1,2}:\d{2}(?::\d{2})?(?:\s?[AaPp]\.?[Mm]\.?)?\b")
_URL = re.compile(r"\b(?:https?|ftp)://\S+", re.IGNORECASE)
_EMAIL = re.compile(r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b")
_WINDOWS_PATH = re.compile(r"\b[A-Za-z]:\\[^\s,;'\"]*")
_UNC_PATH = re.compile(r"\\\\[^\s,;'\"]+")
_UNIX_PATH = re.compile(r"(?:/[\w.\-]+){2,}/?")
_HEX = re.compile(r"\b0[xX][0-9a-fA-F]+\b")
_SINGLE_QUOTED = re.compile(r"'[^']{1,200}'")
_DOUBLE_QUOTED = re.compile(r"\"[^\"]{1,200}\"")
_NUMBER = re.compile(r"(?<![\w<])[+-]?\d[\d,]*(?:\.\d+)?(?![\w>])")
_WHITESPACE = re.compile(r"\s+")
# A trailing "at <date>" / "on <date> <time>" clause says WHEN an error
# happened, never WHICH error it is.  Leaving it in splits one error into two
# rows whenever a sender sometimes appends a timestamp and sometimes does not
# (observed in tests/fixtures: "...timeout: case ID 88421 at 2026-09-28 14:02"
# against a bare "...timeout: case ID 10022").  Removed AFTER placeholders are
# substituted, so only an entirely variable tail can match.
_TEMPORAL_TAIL = re.compile(
    r"(?:[,;:\-]\s*)?"
    r"(?:\b(?:at|on|around|as\s+of|from|dated)\s+)?"
    r"(?:<date>|<time>)(?:\s+(?:<date>|<time>))*"
    r"\s*$",
    re.IGNORECASE,
)
_TRAILING_PUNCT = re.compile(r"[\s.,;:!\-]+$")
_LEADING_PUNCT = re.compile(r"^[\s.,;:\-]+")

# Applied in this order.  Specific patterns must run before general ones or a
# bare number rule would eat the digits out of a date before it is recognised.
_NORMALISERS: tuple[tuple[re.Pattern[str], str], ...] = (
    (_GUID, "<guid>"),
    (_ISO_DATETIME, "<date>"),
    (_URL, "<url>"),
    (_EMAIL, "<email>"),
    (_WINDOWS_PATH, "<path>"),
    (_UNC_PATH, "<path>"),
    (_DATE, "<date>"),
    (_TIME, "<time>"),
    (_HEX, "<hex>"),
    (_UNIX_PATH, "<path>"),
)

# A quoted span this large is the whole message rather than a variable value;
# replacing it would collapse every quoted error into one signature.
_MAX_QUOTED_SHARE = 0.8


def normalize(text: str) -> str:
    """Reduce an error text to its stable signature form.

    Variable parts - timestamps, numbers, paths, GUIDs, quoted values - become
    placeholders, so two occurrences of the same failure produce the same
    string.  The result is lowercased and whitespace-collapsed.

    Args:
        text: Error text, normally already passed through extract_ids().

    Returns:
        The signature.  Empty string if nothing meaningful remains.
    """
    if not text:
        return ""

    working = text
    for pattern, placeholder in _NORMALISERS:
        working = pattern.sub(placeholder, working)

    working = _replace_quoted(working, _SINGLE_QUOTED)
    working = _replace_quoted(working, _DOUBLE_QUOTED)
    working = _NUMBER.sub("<n>", working)

    working = working.lower()
    working = _WHITESPACE.sub(" ", working)
    working = _strip_temporal_tail(working)
    working = _LEADING_PUNCT.sub("", working)
    working = _TRAILING_PUNCT.sub("", working)
    return working.strip()


def _strip_temporal_tail(text: str) -> str:
    """Drop a trailing clause that is nothing but a date or time placeholder.

    Applied repeatedly, because "... at <date> <time>" can leave another
    placeholder exposed once the first is removed.  A tail is only removed
    while something else remains - an error whose entire text is a timestamp
    keeps it rather than becoming an empty signature.

    Args:
        text: The normalized, lowercased text.

    Returns:
        The text without its trailing temporal clause.
    """
    working = text
    while True:
        stripped = _TEMPORAL_TAIL.sub("", working).strip()
        if stripped == working or not stripped:
            return working
        working = stripped


def _replace_quoted(text: str, pattern: re.Pattern[str]) -> str:
    """Replace quoted spans with <q>, unless one span is most of the text.

    A quoted value inside a sentence is a variable and should generalise.  A
    message that is entirely quoted is not - collapsing it to "<q>" would make
    every such error look identical.

    Args:
        text: Text being normalised.
        pattern: The quoting pattern to apply.

    Returns:
        The text with qualifying quoted spans replaced.
    """
    if not text:
        return text

    def _maybe(match: re.Match[str]) -> str:
        if len(match.group(0)) / len(text) >= _MAX_QUOTED_SHARE:
            return match.group(0)
        return "<q>"

    return pattern.sub(_maybe, text)


def _tidy(text: str) -> str:
    """Collapse whitespace and trim stray punctuation left by a removal.

    Args:
        text: Text to tidy.

    Returns:
        The tidied text.
    """
    tidied = _WHITESPACE.sub(" ", text)
    tidied = re.sub(r"\s+([,;:.])", r"\1", tidied)
    tidied = re.sub(r"([(\[])\s+", r"\1", tidied)
    tidied = re.sub(r"\s+([)\]])", r"\1", tidied)
    tidied = re.sub(r"[,;:]\s*(?=[,;:])", "", tidied)
    tidied = _LEADING_PUNCT.sub("", tidied)
    tidied = _TRAILING_PUNCT.sub("", tidied)
    return tidied.strip()


def signature_of(
    error_text: str,
    id_labels: Iterable[str] = DEFAULT_ID_LABELS,
) -> tuple[str, list[str]]:
    """Turn raw error text into a signature plus the identifiers removed.

    Args:
        error_text: The text following a configured marker.
        id_labels: Identifier label words from Config.

    Returns:
        A tuple of (signature, identifier values).
    """
    without_ids, ids = extract_ids(error_text, id_labels)
    return normalize(without_ids), ids


# --------------------------------------------------------------------------
# Fuzzy grouping
# --------------------------------------------------------------------------

def common_wording(left: str, right: str) -> str:
    """Return only the words the two messages share, in order.

    THE GROUP'S NAME, as the operator described it: "we take the message
    Process Failed for case as requestor lock. we remove things which are not
    common." Whatever differs between occurrences — case numbers, sequence
    numbers, operator names, the contents of XML tags — is by definition not
    part of what the error IS, so it falls out on its own without anybody
    listing what to strip.

    Measured on the operator's own examples:
        "No Active Case Case Number 2 sequence 32"
        "No Active Case Case 3 something else"            -> "No Active Case"
        "No Active Case <something>asasas</something>..."

        "Process Failed for Case 123 ... as requestor lock"
        "Process Failed for Case 45435 ... as requestor lock"
            -> "Process Failed for Case as requestor lock"

    Applied one occurrence at a time, so the name narrows as evidence arrives
    and never has to be recomputed from history.

    Args:
        left: The wording so far (or the first message seen).
        right: A new occurrence of the same error.

    Returns:
        The shared wording. Can be empty if the two share nothing, which the
        caller must treat as "no usable name" rather than as a name.
    """
    if not left:
        return (right or "").strip()
    if not right:
        return left.strip()

    first, second = left.split(), right.split()
    kept: list[str] = []
    # autojunk=False for the same reason as similarity(): a long message is
    # mostly "popular" tokens, and the heuristic would discard the very words
    # the two messages have in common.
    matcher = SequenceMatcher(None, first, second, autojunk=False)
    for block in matcher.get_matching_blocks():
        kept.extend(first[block.a:block.a + block.size])
    return " ".join(kept).strip()


def similarity(left: str, right: str) -> float:
    """Return how alike two signatures are, from 0.0 to 1.0.

    Uses difflib.SequenceMatcher, which is in the Python standard library -
    no model, no network, no randomness.

    Args:
        left: First signature.
        right: Second signature.

    Returns:
        The similarity ratio.
    """
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    # autojunk=False IS NOT OPTIONAL HERE, and leaving it at difflib's default
    # was a real defect. The heuristic treats any character occurring in more
    # than 1% of a sequence longer than 200 characters as junk and ignores it.
    # An error message carrying an XML payload is well over 200 characters and
    # is mostly such characters, so the ratio collapses: two occurrences of ONE
    # error measured 0.775 where the true ratio is 0.958, fell under the 0.90
    # threshold, and each occurrence opened its own ledger row for ever.
    # Short messages were unaffected, which is exactly why it went unnoticed.
    return SequenceMatcher(None, left, right, autojunk=False).ratio()


def best_match(
    signature: str,
    known: Iterable[str],
    threshold: float,
) -> tuple[str | None, float]:
    """Find the known signature most like this one.

    An exact match always wins and short-circuits.  Otherwise every known
    signature is scored and the highest is returned, provided it clears the
    threshold.

    Args:
        signature: The incoming signature.
        known: Signatures already in the ledger (and ones created earlier in
            this same run, so a new error seen five times makes one row).
        threshold: Minimum ratio, from Config.

    Returns:
        A tuple of (matched signature or None, the ratio achieved).  The ratio
        is returned even when it falls short, so --explain can show near misses.
    """
    if not signature:
        return None, 0.0

    known_list = list(known)
    if signature in known_list:
        return signature, 1.0

    best_signature: str | None = None
    best_ratio = 0.0
    # autojunk=False for the same reason as in similarity() — see the note
    # there. The two must agree, or --explain would print a ratio the matcher
    # did not act on.
    matcher = SequenceMatcher(None, autojunk=False)
    matcher.set_seq2(signature)

    for candidate in known_list:
        matcher.set_seq1(candidate)
        # Cheap upper bounds first - skip candidates that cannot possibly win.
        if matcher.real_quick_ratio() < best_ratio:
            continue
        if matcher.quick_ratio() < best_ratio:
            continue
        ratio = matcher.ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_signature = candidate

    if best_signature is not None and best_ratio >= threshold:
        return best_signature, best_ratio
    return None, best_ratio
