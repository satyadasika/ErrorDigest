# ErrorDigest

Reads error emails from Outlook folders, works out which errors are *the same
error*, keeps a permanent count in one Excel file, and writes you a draft email
summarising what happened. You review the draft and click Send.

**There is no AI in this program.** Grouping is regular expressions plus
`difflib`, both part of Python itself. The same input always gives the same
output, and every decision can be explained.

**It cannot send email.** Drafts are created with `Save()` and `Display()`.
`Send()` appears nowhere in the code, and `tests/test_never_sends.py` reads the
source of every file to prove it — twice, by two different methods.

---

## Setup on Windows (once, about 3 minutes)

You need **Python 3.10 or newer** and **Outlook desktop** (the classic app, not
the web version). Nothing else is installed, nothing runs in the background,
and nothing costs anything.

1. If you do not have Python: https://www.python.org/downloads/windows/ —
   on the first screen **tick "Add python.exe to PATH"**.
2. Copy this folder anywhere, e.g. `C:\Tools\ErrorDigest`.
3. Double-click **`setup.bat`**.

That creates a private environment in `.venv\`, installs the two packages
(`openpyxl` and `pywin32`), and creates `error_digest.xlsx`.

4. Open **`error_digest.xlsx`**, go to the **Config** sheet, and fill in:

   | Setting | Set it to |
   |---|---|
   | `folders` | The Outlook folder path, exactly as the folder pane shows it, e.g. `Inbox\Alerts`. Several folders go one per line — press **Alt+Enter** inside the cell for a new line. |
   | `draft_to` | Who the digest should be addressed to. Several addresses separated by `;`. |

5. Close Excel (the script cannot write the file while it is open), then
   double-click **`run.bat`**.

---

## Running it from VSCode

Open this folder in VSCode (`File → Open Folder`). Install the **Python**
extension if prompted.

Press **F5** and pick one of the ready-made configurations:

| Configuration | What it does |
|---|---|
| **Run digest (creates the draft)** | The normal run. |
| **Run digest - no draft** | Updates the ledger only. |
| **Dry run** | Reports what it *would* do. Writes nothing. |
| **Test offline (fixtures, no Outlook needed)** | Runs the whole pipeline against the sample emails in `tests/fixtures`. Works even with Outlook closed. |
| **Create the workbook (first time)** | Same as `--init`. |

If VSCode says it cannot find the interpreter: `Ctrl+Shift+P` →
**Python: Select Interpreter** → choose `.\.venv\Scripts\python.exe`.

To install **ExcelViewer** (recommended in `.vscode/extensions.json`) you can
read the workbook inside VSCode without opening Excel — handy, because the
script cannot save while Excel has the file open.

---

## How it decides two errors are the same

```
Error Message: Payment gateway timeout: case ID 88421 at 2026-09-28 14:02
Error Message: Payment gateway timeout: case ID 99110 at 2026-09-29 09:14
Error Message:
Payment gateway timeout: case ID 10022
                            |
      1. cut off quoted history (forwarded / replied text)
      2. find the marker from the Keywords sheet, take what follows
      3. pull OUT the identifiers      -> 88421, 99110, 10022 recorded
      4. replace variable parts        -> dates, numbers, paths, GUIDs
      5. drop a trailing timestamp     -> "at <date>" says when, not which
                            |
                 signature: "payment gateway timeout"
                            |
      6. match the ledger: exact, else similarity >= fuzzy_threshold
                            |
          ONE row. Total 3. First Found 2026-09-28. IDs: 88421, 99110, 10022
```

Step 6 is what catches wording that drifts:

```
"connection timeout on host srv-<n>"      \
"connection timed out on host srv-<n>"    /  similarity 0.97  ->  grouped
```

Run with `--explain` to see the similarity score for every error, including
the near misses — that is how you pick the right `fuzzy_threshold` from your
own mail instead of guessing.

---

## The workbook

Everything is configured here. You never edit code.

### `Config`
Settings, each with a plain-English explanation in the third column. The one to
know about is **`last_run`**: the script writes it after each successful run and
reads only email that arrived after it. **Clear that cell to re-read history.**

### `Keywords`
Your error-start markers, one per row, each with **Enabled** `Y`/`N`:
`Error Message`, `Exception`, `Error`, `Failure`, `Failed`, `FATAL`,
and `Warning` (off by default).

Add your own rows at any time. Matching is case-insensitive and whole-word, so
`Error` never fires inside `Errors` or `Terror`. The longest marker wins, so a
line starting `Error Message:` is attributed to `Error Message`, not `Error`.

### `Errors` — the ledger
One row per distinct error.

| Column | Who owns it |
|---|---|
| Signature, Sample Message, Total Count, Count This Run, First Found, Last Seen, Case IDs Seen, Source Folders, Last Updated | the script |
| **Status, Owner, Notes** | **you** — tinted yellow, and the script never writes them |

**First Found is written once and never changed.** That is the "we have seen
this before, since 14 August" date the digest reports.

### `Occurrences` — the history
One row per individual error instance: when, which folder, which email, which
identifiers. This is also the de-duplication record — the Entry ID column is
what makes running twice harmless. Rows older than
`occurrence_retention_days` (default 180) are pruned so the file does not grow
forever. The `Errors` ledger is **never** pruned.

---

## Running it on a schedule

Windows Task Scheduler (this is Windows, so not cron):

- **Program:** `C:\Tools\ErrorDigest\run_silent.bat`
- **Start in:** `C:\Tools\ErrorDigest`
- **Trigger:** whatever suits, e.g. hourly on weekdays
- **Settings:** choose **"Run only when the user is logged on"** — Outlook
  automation needs a real desktop session.

`run_silent.bat` collects into the workbook and creates no draft, appending to
`digest_log.txt`. When you want the email, double-click `run.bat`, or run
`error_digest.py --draft-only` to build a draft from what the ledger already
holds without reading any more mail.

---

## Command line

```
python error_digest.py                     normal run: read, update, draft
python error_digest.py --no-draft          ledger only, for scheduled runs
python error_digest.py --draft-only        draft from the ledger, read nothing
python error_digest.py --dry-run --explain report only, change nothing
python error_digest.py --since 2026-09-01  override the window start
python error_digest.py --init              create the workbook, then stop
python error_digest.py --source fixtures --fixtures tests\fixtures --since 2026-09-01
                                           offline test, no Outlook required
```

---

## When something goes wrong

| What you see | What it means |
|---|---|
| `it is open in Excel` | Close the workbook and run again. **Nothing was lost** — `last_run` is only advanced after a successful save, so the same emails are read next time. |
| `No folder 'X' under 'Inbox'` | The path is wrong. The message lists the folders that *do* exist there — usually enough to spot the typo. Copy the name exactly as Outlook's folder pane shows it. |
| `pywin32 is not installed` | Run `setup.bat`, or `pip install pywin32`. |
| `Could not start Outlook` | Open Outlook, then run again. |
| `Could not create the Outlook draft` | The ledger is still saved. The digest is written to `last_digest.html` — open it in a browser. |
| Two errors that should be one | Lower `fuzzy_threshold` (e.g. 0.85). Use `--explain` to see what score they actually got. |
| One row holding two different errors | Raise `fuzzy_threshold` (e.g. 0.95). |
| An error is not picked up at all | Its line does not start with any enabled marker. Add the marker to the `Keywords` sheet. |

---

## Tests

```
.venv\Scripts\python.exe -m pip install pytest
.venv\Scripts\python.exe -m pytest tests -q
```

75 tests, and they run on any operating system — the Outlook layer is isolated
in `outlook_win.py` and swapped for `fixture_source.py`, so parsing, grouping,
the Excel ledger and the digest email are all exercised for real.

---

## Files

| File | What it is |
|---|---|
| `error_digest.py` | The program. This is what you run. |
| `parsing.py` | Marker matching, ID stripping, normalising, fuzzy grouping. No I/O. |
| `workbook.py` | Reads and writes the Excel file. Protects your Status/Owner/Notes. |
| `outlook_win.py` | The only Windows-specific file. Reads folders, creates the draft. |
| `fixture_source.py` | Stand-in for Outlook, used by the tests and `--source fixtures`. |
| `message.py` | The shared email record both sources produce. |
| `tests/` | 75 tests, including the tripwire that proves nothing can send. |

---

## Known limits, stated rather than hidden

- **An email with no errors is not recorded**, so if you widen the window with
  `--since` it is re-read. It produces nothing, so this is harmless — but it is
  why the "processed" count can be higher than you expect.
- **The error text is taken from one line.** An error whose message wraps over
  several lines contributes only its first line to the signature. In practice
  machine-generated alerts put the message on one line; if yours do not, say so
  and the rule can change.
- **Removing an identifier can leave a dangling word**, e.g.
  `Invalid routing key for ticket #7781` becomes the signature
  `invalid routing key for`. Grouping is still correct and consistent; the
  `Sample Message` column always shows the real, verbatim text.
- **Outlook must be running** for a scheduled run to work, and the session must
  be logged on.
