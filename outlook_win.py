"""
Outlook for Windows adapter - the only Windows-specific file in the project.

Everything here goes through pywin32 COM against the Outlook desktop
application that is already running and already signed in, so there are no
credentials, no API keys and no network configuration.

THE ONE RULE THIS FILE EXISTS TO KEEP: it never sends anything.  `.Send()`
does not appear in this project, and tests/test_never_sends.py reads the source
of every file to prove it.  Drafts are created with `.Save()` (which puts them
in Drafts) and `.Display()` (which opens the window), and the person clicks
Send themselves.
"""

from __future__ import annotations

import datetime as dt
from typing import Iterable

from message import EmailMessage, to_naive_local

# Outlook enumeration values.  Spelled out rather than inlined as magic numbers.
OL_FOLDER_INBOX = 6          # olFolderInbox
OL_MAIL_ITEM = 0             # olMailItem
OL_CLASS_MAIL = 43           # olMail


class OutlookUnavailable(RuntimeError):
    """Raised when Outlook or pywin32 cannot be reached.

    Carries a message that says what to do about it, because this is the error
    a person is most likely to meet on their first run.
    """


def _application():
    """Attach to the running Outlook application.

    Returns:
        The Outlook.Application COM object.

    Raises:
        OutlookUnavailable: If pywin32 is not installed or Outlook will not
            start - for example when running on macOS or Linux.
    """
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise OutlookUnavailable(
            "pywin32 is not installed, or this is not Windows. "
            "On Windows run:  pip install pywin32"
        ) from exc

    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:  # pragma: no cover - needs a real Windows box
        raise OutlookUnavailable(
            f"Could not start Outlook ({exc}). Open Outlook and try again."
        ) from exc


def resolve_folder(namespace, folder_path: str):
    """Walk a backslash-separated folder path to the folder object.

    The first segment may be "Inbox" (the default inbox of the default
    account) or the name of a mailbox or PST store as it appears in Outlook's
    folder pane, which is what makes shared and secondary mailboxes reachable.

    Args:
        namespace: The MAPI namespace from the Outlook application.
        folder_path: For example "Inbox\\Alerts\\Prod".

    Returns:
        The Outlook folder object.

    Raises:
        OutlookUnavailable: If any segment of the path does not exist.  The
            message lists the folders that DO exist at the point it failed,
            which is almost always enough to spot a typo.
    """
    segments = [part.strip() for part in folder_path.replace("/", "\\").split("\\") if part.strip()]
    if not segments:
        raise OutlookUnavailable("Empty folder path in Config.")

    head, *rest = segments
    if head.lower() == "inbox":
        current = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    else:
        try:
            current = namespace.Folders(head)
        except Exception as exc:
            available = _child_names(namespace.Folders)
            raise OutlookUnavailable(
                f"No mailbox or folder named '{head}'. Top level holds: {available}"
            ) from exc

    for segment in rest:
        try:
            current = current.Folders(segment)
        except Exception as exc:
            available = _child_names(current.Folders)
            raise OutlookUnavailable(
                f"No folder '{segment}' under '{current.Name}'. It holds: {available}"
            ) from exc

    return current


def _child_names(folders) -> str:
    """List child folder names for an error message.

    Args:
        folders: An Outlook Folders collection.

    Returns:
        A comma-separated list, or "(none)".
    """
    try:
        names = [folders.Item(i).Name for i in range(1, folders.Count + 1)]
    except Exception:  # pragma: no cover - defensive, only in error paths
        return "(could not be listed)"
    return ", ".join(names) if names else "(none)"


def list_folders(max_depth: int = 4) -> list[str]:
    """Return every mail folder path, the way the Config sheet wants them written.

    Built because the commonest failure on a first run is a folder path that is
    nearly right — "Inbox\\Alerts" when the real tree has "Inbox\\Alerts\\Prod",
    or a shared mailbox whose display name differs from what the person calls it.
    Guessing costs several failed runs; reading the real list costs one command.

    Args:
        max_depth: How deep to walk. The default covers any realistic alert tree
            and stops a pathological mailbox from printing thousands of lines.

    Returns:
        Folder paths, e.g. ["Inbox", "Inbox\\Alerts", "Inbox\\Alerts\\Prod"].

    Raises:
        OutlookUnavailable: If Outlook cannot be reached.
    """
    application = _application()
    namespace = application.GetNamespace("MAPI")
    found: list[str] = []

    def walk(folder, prefix: str, depth: int) -> None:
        """Add this folder's path, then recurse into its children."""
        found.append(prefix)
        if depth >= max_depth:
            return
        try:
            children = folder.Folders
            for index in range(1, children.Count + 1):
                child = children.Item(index)
                walk(child, f"{prefix}\\{child.Name}", depth + 1)
        except Exception:  # pragma: no cover - a folder that refuses to enumerate
            return

    # The default inbox first, because that is what nearly every path starts with.
    try:
        walk(namespace.GetDefaultFolder(OL_FOLDER_INBOX), "Inbox", 1)
    except Exception:  # pragma: no cover - no default inbox on this profile
        pass

    # Then every store (other mailboxes, shared mailboxes, PST files), so a path
    # that does NOT start at Inbox is discoverable too.
    try:
        stores = namespace.Folders
        for index in range(1, stores.Count + 1):
            store = stores.Item(index)
            walk(store, store.Name, 1)
    except Exception:  # pragma: no cover
        pass

    seen: set[str] = set()
    unique = []
    for path in found:
        if path not in seen:
            seen.add(path)
            unique.append(path)
    return unique


def fetch_messages(folder_paths: Iterable[str], since: dt.datetime) -> list[EmailMessage]:
    """Read every mail item received at or after a moment, from each folder.

    Outlook's Restrict() is tried first because it filters inside Outlook and
    is far faster on a large folder.  Its date syntax is locale-sensitive, so
    a failure falls back to filtering in Python rather than returning nothing -
    slower, but it cannot silently skip mail.

    Args:
        folder_paths: Folder paths to scan.
        since: Only messages received at or after this local moment.

    Returns:
        Every matching message, oldest first.

    Raises:
        OutlookUnavailable: If Outlook cannot be reached or a folder is wrong.
    """
    application = _application()
    namespace = application.GetNamespace("MAPI")
    collected: list[EmailMessage] = []

    for folder_path in folder_paths:
        folder = resolve_folder(namespace, folder_path)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        restricted = items
        try:
            stamp = since.strftime("%m/%d/%Y %I:%M %p")
            restricted = items.Restrict(f"[ReceivedTime] >= '{stamp}'")
        except Exception:
            restricted = items  # locale mismatch - filter in Python below

        for item in restricted:
            try:
                if int(getattr(item, "Class", OL_CLASS_MAIL)) != OL_CLASS_MAIL:
                    continue
                received = to_naive_local(item.ReceivedTime)
            except (AttributeError, TypeError, ValueError):
                continue  # meeting requests, reports and similar - no body to read

            if received < since:
                continue

            collected.append(EmailMessage(
                entry_id=str(getattr(item, "EntryID", "") or ""),
                folder_path=folder_path,
                subject=str(getattr(item, "Subject", "") or ""),
                sender=_sender_of(item),
                received=received,
                body=str(getattr(item, "Body", "") or ""),
            ))

    collected.sort(key=lambda message: message.received)
    return collected


def _sender_of(item) -> str:
    """Read the sender's display name, falling back to the address.

    Args:
        item: An Outlook mail item.

    Returns:
        The sender, or "" when neither property is readable.
    """
    for attribute in ("SenderName", "SenderEmailAddress"):
        try:
            value = getattr(item, attribute, None)
            if value:
                return str(value)
        except Exception:  # pragma: no cover - some items refuse these
            continue
    return ""


def create_draft(
    to: str,
    cc: str,
    subject: str,
    html_body: str,
    display: bool = True,
) -> None:
    """Create an UNSENT draft in Outlook and open it for review.

    The draft is saved to the Drafts folder and shown on screen.  It is never
    sent - the person reads it, edits it if they want, and clicks Send.

    Args:
        to: Recipients, separated by ';'.
        cc: Cc recipients, separated by ';'.
        subject: Subject line.
        html_body: The digest table as HTML.
        display: Open the draft window.  False still saves it to Drafts.

    Raises:
        OutlookUnavailable: If Outlook cannot be reached.
    """
    application = _application()
    mail = application.CreateItem(OL_MAIL_ITEM)
    mail.To = to or ""
    mail.CC = cc or ""
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Save()          # lands in Drafts
    if display:
        mail.Display(False)   # modeless, so the script can exit
