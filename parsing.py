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
_MARKER_SEPARATOR = re.compile(r"^\s*[:\-–—>|]*\s*")


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

def find_error_texts(
    body: str,
    keywords: Sequence[Keyword],
    max_errors: int = 0,
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

    for index, line in enumerate(lines):
        for marker, pattern in compiled:
            match = pattern.search(line)
            if not match:
                continue

            remainder = line[match.end():]
            remainder = _MARKER_SEPARATOR.sub("", remainder).strip()

            if not remainder:
                remainder = _next_non_empty(lines, index + 1)

            if remainder:
                found.append((marker, remainder))
                if max_errors and len(found) >= max_errors:
                    return found
            break  # one marker per line - the longest that matched

    return found


def _next_non_empty(lines: Sequence[str], start: int) -> str:
    """Return the first non-blank line at or after an index.

    Args:
        lines: All body lines.
        start: Index to begin looking from.

    Returns:
        The stripped line, or "" if every remaining line is blank.
    """
    for line in lines[start:]:
        stripped = line.strip()
        if stripped:
            return stripped
    return ""


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
    return SequenceMatcher(None, left, right).ratio()


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
    matcher = SequenceMatcher(None)
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
