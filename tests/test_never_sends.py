"""
The guarantee that matters most: this program cannot send email.

Drafts are created with Save() and Display().  Send() is never called.  These
tests read the actual source of every Python file in the project, so the
guarantee cannot be broken later by an edit that looks harmless - including an
edit made by an AI assistant.
"""

import ast
import re
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SOURCE_FILES = sorted(PROJECT_ROOT.glob("*.py"))


def test_there_are_source_files_to_check():
    """A glob that silently matched nothing would make every test below vacuous."""
    assert len(SOURCE_FILES) >= 5


def _code_only(path: Path) -> str:
    """Return a file's source with comments and string literals blanked out.

    Written because the first version of this test failed on outlook_win.py's
    own docstring, which says ".Send()" while explaining that it never calls
    it.  A tripwire that cannot tell code from prose gets switched off, and a
    switched-off tripwire protects nothing.

    Args:
        path: The Python file to read.

    Returns:
        The source with every comment and string replaced by blanks of the
        same length, so line numbers still line up.
    """
    import io
    import tokenize

    source = path.read_text(encoding="utf-8")
    out = list(source)
    readline = io.StringIO(source).readline
    offsets = [0]
    for line in source.splitlines(keepends=True):
        offsets.append(offsets[-1] + len(line))

    for token in tokenize.generate_tokens(readline):
        if token.type not in (tokenize.COMMENT, tokenize.STRING):
            continue
        start = offsets[token.start[0] - 1] + token.start[1]
        end = offsets[token.end[0] - 1] + token.end[1]
        for index in range(start, min(end, len(out))):
            if out[index] != "\n":
                out[index] = " "
    return "".join(out)


@pytest.mark.parametrize("path", SOURCE_FILES, ids=lambda p: p.name)
def test_no_send_call_anywhere(path):
    """No file may contain a .Send( call in executable code."""
    offenders = [
        f"line {number}: {line.strip()}"
        for number, line in enumerate(_code_only(path).splitlines(), start=1)
        if re.search(r"\.\s*Send\s*\(", line)
    ]
    assert not offenders, f"{path.name} contains a Send() call: {offenders}"


@pytest.mark.parametrize("path", SOURCE_FILES, ids=lambda p: p.name)
def test_no_send_attribute_access_in_the_parsed_tree(path):
    """
    A second, independent check against the parsed syntax tree.

    The regex above reads text and could be defeated by unusual spacing or a
    string built at runtime; this one reads the AST.  Two different methods
    agreeing is the point - one of them alone is a promise, both is a fact.
    """
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
            assert node.func.attr not in {"Send", "send"}, (
                f"{path.name} line {node.lineno} calls .{node.func.attr}()"
            )


def test_the_draft_function_saves_and_displays():
    """create_draft must actually create a draft - the opposite failure."""
    source = (PROJECT_ROOT / "outlook_win.py").read_text(encoding="utf-8")
    body = source.split("def create_draft")[1]
    assert "mail.Save()" in body
    assert "mail.Display(" in body


def test_the_tripwire_itself_works():
    """
    Prove the check fails on a file that DOES send.

    A guard that has never seen the thing it guards against is not a guard.
    """
    probe = "mail = app.CreateItem(0)\nmail.Send()\n"
    assert re.search(r"\.\s*Send\s*\(", probe)
    # and the comment/string blanking must NOT hide a genuine call
    import tempfile
    with tempfile.NamedTemporaryFile("w", suffix=".py", delete=False) as handle:
        handle.write('"""A docstring mentioning .Send() harmlessly."""\n' + probe)
        probe_path = Path(handle.name)
    blanked = _code_only(probe_path)
    assert "Send(" in blanked, "a real call was wrongly blanked out"
    assert "docstring" not in blanked, "prose was not blanked out"
    probe_path.unlink()
    tree = ast.parse(probe)
    found = [
        node.func.attr for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
    ]
    assert "Send" in found
