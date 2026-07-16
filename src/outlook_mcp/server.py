"""Outlook MCP server backed by COM automation (win32com).

Talks directly to a locally installed, classic Outlook desktop client via the
Outlook Object Model. No Microsoft Graph, no Azure app registration, no OAuth:
whatever profile Outlook is signed into is what this server operates on.

Requirements:
  * Windows
  * Classic Outlook desktop (the "new Outlook" has no COM interface)
  * pywin32

Every tool opens its own short-lived COM session (CoInitialize + Dispatch)
under a global lock. Outlook.Application is a singleton, so Dispatch simply
attaches to the running instance (or starts one).
"""

from __future__ import annotations

import sys
import threading
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterator

from mcp.server.fastmcp import FastMCP

try:
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover - non-Windows platforms
    pythoncom = None
    win32com = None

mcp = FastMCP("outlook")

# Outlook Object Model constants (OlDefaultFolders / OlItemType / etc.)
FOLDER_INBOX = 6
FOLDER_SENT = 5
FOLDER_DRAFTS = 16
FOLDER_OUTBOX = 4
FOLDER_DELETED = 3
FOLDER_JUNK = 23
FOLDER_CALENDAR = 9
FOLDER_CONTACTS = 10

ITEM_MAIL = 0
ITEM_APPOINTMENT = 1

MAIL_ITEM_CLASS = 43  # olMail

# OlBusyStatus
BUSY_FREE = 0
BUSY_TENTATIVE = 1
BUSY_BUSY = 2
BUSY_OUT_OF_OFFICE = 3
BUSY_WORKING_ELSEWHERE = 4

# Recipient.FreeBusy() with CompleteFormat=True returns one char per time slot.
FREEBUSY_LEGEND = {
    "0": "Free",
    "1": "Tentative",
    "2": "Busy",
    "3": "Out of Office",
    "4": "Working Elsewhere",
}
FREEBUSY_FREE = "0"

WELL_KNOWN_FOLDERS = {
    "inbox": FOLDER_INBOX,
    "sent items": FOLDER_SENT,
    "sent": FOLDER_SENT,
    "drafts": FOLDER_DRAFTS,
    "outbox": FOLDER_OUTBOX,
    "deleted items": FOLDER_DELETED,
    "trash": FOLDER_DELETED,
    "junk email": FOLDER_JUNK,
    "junk": FOLDER_JUNK,
    "spam": FOLDER_JUNK,
    "calendar": FOLDER_CALENDAR,
    "contacts": FOLDER_CONTACTS,
}

# COM objects live in a single-threaded apartment; FastMCP runs sync tools on
# worker threads, so serialize all Outlook access and init COM per call.
_COM_LOCK = threading.Lock()


@contextmanager
def outlook_session() -> Iterator[Any]:
    """Yield the MAPI namespace of a live Outlook instance."""
    if pythoncom is None:
        raise RuntimeError(
            "pywin32 is not available. This server requires Windows with "
            "classic Outlook desktop installed."
        )
    with _COM_LOCK:
        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            yield app.GetNamespace("MAPI")
        finally:
            pythoncom.CoUninitialize()


def resolve_folder(ns: Any, folder: str) -> Any:
    """Resolve a folder by well-known name or slash-separated path.

    Examples: "Inbox", "Sent Items", "Inbox/Receipts",
    "someone@company.com/Inbox" (a specific store).
    """
    parts = [p for p in folder.replace("\\", "/").split("/") if p]
    if not parts:
        raise ValueError("Empty folder name")

    head = parts[0].lower()
    if head in WELL_KNOWN_FOLDERS:
        current = ns.GetDefaultFolder(WELL_KNOWN_FOLDERS[head])
        rest = parts[1:]
    else:
        # Try to match a store (account) root by display name.
        current = None
        for store_folder in ns.Folders:
            if store_folder.Name.lower() == head:
                current = store_folder
                break
        if current is None:
            # Fall back: top-level folder inside the default store.
            root = ns.GetDefaultFolder(FOLDER_INBOX).Parent
            for sub in root.Folders:
                if sub.Name.lower() == head:
                    current = sub
                    break
        if current is None:
            raise ValueError(f"Folder not found: {parts[0]!r}")
        rest = parts[1:]

    for part in rest:
        match = None
        for sub in current.Folders:
            if sub.Name.lower() == part.lower():
                match = sub
                break
        if match is None:
            raise ValueError(f"Subfolder {part!r} not found under {current.Name!r}")
        current = match
    return current


def get_item(ns: Any, entry_id: str) -> Any:
    try:
        return ns.GetItemFromID(entry_id)
    except Exception as exc:
        raise ValueError(f"No item found for entry_id {entry_id!r}: {exc}") from exc


def sender_address(item: Any) -> str:
    """Best-effort SMTP address of the sender (Exchange hides it behind EX)."""
    try:
        if item.SenderEmailType == "EX":
            exch = item.Sender.GetExchangeUser()
            if exch is not None:
                return exch.PrimarySmtpAddress
        return item.SenderEmailAddress or ""
    except Exception:
        return getattr(item, "SenderEmailAddress", "") or ""


def fmt_dt(value: Any) -> str:
    try:
        return value.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(value)


def mail_summary(item: Any) -> dict[str, Any]:
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "from": f"{item.SenderName} <{sender_address(item)}>",
        "to": item.To or "",
        "received": fmt_dt(item.ReceivedTime),
        "unread": bool(item.UnRead),
        "has_attachments": item.Attachments.Count > 0,
        "preview": (item.Body or "")[:200].replace("\r\n", " ").strip(),
    }


def render_mail_list(items: list[dict[str, Any]]) -> str:
    if not items:
        return "No emails found."
    lines = []
    for i, m in enumerate(items, 1):
        flags = []
        if m["unread"]:
            flags.append("UNREAD")
        if m["has_attachments"]:
            flags.append("ATTACHMENTS")
        flag_str = f" [{', '.join(flags)}]" if flags else ""
        lines.append(
            f"{i}. {m['subject']}{flag_str}\n"
            f"   From: {m['from']}  |  Received: {m['received']}\n"
            f"   Preview: {m['preview']}\n"
            f"   entry_id: {m['entry_id']}"
        )
    return "\n\n".join(lines)


def dasl_escape(text: str) -> str:
    return text.replace("'", "''").replace("%", "[%]")


def parse_when(value: str) -> datetime:
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise ValueError(
        f"Could not parse datetime {value!r}. Use 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD'."
    )


def restrict_date(dt: datetime) -> str:
    """Format a datetime the way Outlook's Restrict() expects."""
    return dt.strftime("%m/%d/%Y %I:%M %p")


def parse_date_only(value: str) -> datetime:
    """Parse a YYYY-MM-DD (or YYYY-MM-DD HH:MM) string to midnight of that day."""
    dt = parse_when(value)
    return dt.replace(hour=0, minute=0, second=0, microsecond=0)


def resolve_recipient(ns: Any, address: str) -> Any:
    """Resolve a name or SMTP address to an Outlook Recipient object."""
    recip = ns.CreateRecipient(address)
    recip.Resolve()
    if not recip.Resolved:
        raise ValueError(
            f"Could not resolve {address!r} in the address book / GAL. "
            "Use a full email address or an exact display name."
        )
    return recip


def freebusy_slots(
    recip: Any, start: datetime, min_per_char: int
) -> list[str]:
    """Return the FreeBusy status string for a recipient as a list of slot codes.

    Each entry is a single character from FREEBUSY_LEGEND covering
    ``min_per_char`` minutes, beginning at midnight of ``start``.
    """
    # CompleteFormat=True -> distinguishes free/tentative/busy/OOF.
    raw = recip.FreeBusy(start, min_per_char, True)
    return list(raw or "")


def slot_time(day_start: datetime, index: int, min_per_char: int) -> datetime:
    return day_start + timedelta(minutes=index * min_per_char)


def merge_intervals(intervals: list[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
    """Merge overlapping/adjacent (start, end) intervals."""
    if not intervals:
        return []
    ordered = sorted(intervals)
    merged = [ordered[0]]
    for start, end in ordered[1:]:
        last_start, last_end = merged[-1]
        if start <= last_end:
            merged[-1] = (last_start, max(last_end, end))
        else:
            merged.append((start, end))
    return merged


# ---------------------------------------------------------------------------
# Folder tools
# ---------------------------------------------------------------------------


@mcp.tool()
def list_folders() -> str:
    """List all mail folders across every account/store in Outlook."""
    with outlook_session() as ns:
        lines: list[str] = []

        def walk(folder: Any, depth: int) -> None:
            try:
                count = folder.Items.Count
            except Exception:
                count = "?"
            lines.append(f"{'  ' * depth}- {folder.Name} ({count} items)")
            try:
                for sub in folder.Folders:
                    walk(sub, depth + 1)
            except Exception:
                pass

        for store_folder in ns.Folders:
            lines.append(f"Store: {store_folder.Name}")
            for sub in store_folder.Folders:
                walk(sub, 1)
        return "\n".join(lines) if lines else "No folders found."


# ---------------------------------------------------------------------------
# Email reading tools
# ---------------------------------------------------------------------------


@mcp.tool()
def list_emails(folder: str = "Inbox", count: int = 20, unread_only: bool = False) -> str:
    """List recent emails in a folder, newest first.

    Args:
        folder: Folder name or path, e.g. "Inbox", "Sent Items", "Inbox/Receipts".
        count: Maximum number of emails to return (1-100).
        unread_only: Only return unread emails.
    """
    count = max(1, min(count, 100))
    with outlook_session() as ns:
        items = resolve_folder(ns, folder).Items
        items.Sort("[ReceivedTime]", True)
        if unread_only:
            items = items.Restrict("[UnRead] = True")

        results = []
        for item in items:
            if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                continue
            results.append(mail_summary(item))
            if len(results) >= count:
                break
        return render_mail_list(results)


@mcp.tool()
def get_email(entry_id: str, body_max_chars: int = 8000) -> str:
    """Read a full email by its entry_id (from list_emails/search_emails).

    Args:
        entry_id: The Outlook EntryID of the email.
        body_max_chars: Truncate the body after this many characters.
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        body = (item.Body or "").strip()
        truncated = len(body) > body_max_chars
        attachments = [
            f"{a.FileName} ({a.Size} bytes)" for a in item.Attachments
        ]
        parts = [
            f"Subject: {item.Subject}",
            f"From: {item.SenderName} <{sender_address(item)}>",
            f"To: {item.To or ''}",
        ]
        if item.CC:
            parts.append(f"CC: {item.CC}")
        parts.append(f"Received: {fmt_dt(item.ReceivedTime)}")
        parts.append(f"entry_id: {item.EntryID}")
        if attachments:
            parts.append("Attachments: " + ", ".join(attachments))
        parts.append("")
        parts.append(body[:body_max_chars])
        if truncated:
            parts.append(f"\n[... body truncated at {body_max_chars} chars ...]")
        return "\n".join(parts)


@mcp.tool()
def search_emails(
    query: str,
    folder: str = "Inbox",
    count: int = 20,
    search_in: str = "all",
) -> str:
    """Search emails in a folder by subject, sender, or body text.

    Args:
        query: Text to search for.
        folder: Folder to search in.
        count: Maximum results (1-100).
        search_in: One of "subject", "from", "body", "all".
    """
    count = max(1, min(count, 100))
    q = dasl_escape(query)
    fields = {
        "subject": ["urn:schemas:httpmail:subject"],
        "from": ["urn:schemas:httpmail:fromname", "urn:schemas:httpmail:fromemail"],
        "body": ["urn:schemas:httpmail:textdescription"],
        "all": [
            "urn:schemas:httpmail:subject",
            "urn:schemas:httpmail:fromname",
            "urn:schemas:httpmail:fromemail",
            "urn:schemas:httpmail:textdescription",
        ],
    }
    if search_in not in fields:
        raise ValueError("search_in must be one of: subject, from, body, all")
    clause = " OR ".join(f'"{f}" LIKE \'%{q}%\'' for f in fields[search_in])
    dasl = f"@SQL=({clause})"

    with outlook_session() as ns:
        items = resolve_folder(ns, folder).Items
        items.Sort("[ReceivedTime]", True)
        items = items.Restrict(dasl)
        results = []
        for item in items:
            if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                continue
            results.append(mail_summary(item))
            if len(results) >= count:
                break
        return render_mail_list(results)


@mcp.tool()
def save_attachments(entry_id: str, save_dir: str) -> str:
    """Save all attachments of an email to a local directory.

    Args:
        entry_id: The email's EntryID.
        save_dir: Directory to save into (created if missing).
    """
    target = Path(save_dir).expanduser()
    target.mkdir(parents=True, exist_ok=True)
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        if item.Attachments.Count == 0:
            return "This email has no attachments."
        saved = []
        for att in item.Attachments:
            dest = target / att.FileName
            att.SaveAsFile(str(dest))
            saved.append(str(dest))
        return "Saved attachments:\n" + "\n".join(saved)


# ---------------------------------------------------------------------------
# Email writing tools
# ---------------------------------------------------------------------------


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html: bool = False,
    attachments: list[str] | None = None,
    save_as_draft: bool = False,
) -> str:
    """Send a new email (or save it as a draft) through Outlook.

    Args:
        to: Recipient address(es), separated by semicolons.
        subject: Email subject.
        body: Email body text (or HTML if html=True).
        cc: CC address(es), semicolon-separated.
        bcc: BCC address(es), semicolon-separated.
        html: Treat body as HTML.
        attachments: Absolute paths of files to attach.
        save_as_draft: Save to Drafts instead of sending.
    """
    with outlook_session() as ns:
        app = ns.Application
        mail = app.CreateItem(ITEM_MAIL)
        mail.To = to
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        mail.Subject = subject
        if html:
            mail.HTMLBody = body
        else:
            mail.Body = body
        for path in attachments or []:
            file = Path(path).expanduser()
            if not file.is_file():
                raise ValueError(f"Attachment not found: {path}")
            mail.Attachments.Add(str(file))
        if save_as_draft:
            mail.Save()
            return f"Draft saved: {subject!r} to {to}"
        mail.Send()
        return f"Email sent: {subject!r} to {to}"


@mcp.tool()
def reply_to_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    save_as_draft: bool = False,
) -> str:
    """Reply to an email. The original message is quoted below your reply.

    Args:
        entry_id: EntryID of the email to reply to.
        body: Your reply text.
        reply_all: Reply to all recipients instead of just the sender.
        save_as_draft: Save to Drafts instead of sending.
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        reply = item.ReplyAll() if reply_all else item.Reply()
        reply.Body = body + "\n\n" + reply.Body
        if save_as_draft:
            reply.Save()
            return f"Reply draft saved for: {item.Subject!r}"
        reply.Send()
        return f"Reply sent for: {item.Subject!r}"


@mcp.tool()
def forward_email(
    entry_id: str,
    to: str,
    comment: str = "",
    save_as_draft: bool = False,
) -> str:
    """Forward an email (attachments included).

    Args:
        entry_id: EntryID of the email to forward.
        to: Recipient address(es), semicolon-separated.
        comment: Optional text placed above the forwarded message.
        save_as_draft: Save to Drafts instead of sending.
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        fwd = item.Forward()
        fwd.To = to
        if comment:
            fwd.Body = comment + "\n\n" + fwd.Body
        if save_as_draft:
            fwd.Save()
            return f"Forward draft saved for: {item.Subject!r}"
        fwd.Send()
        return f"Forwarded {item.Subject!r} to {to}"


# ---------------------------------------------------------------------------
# Email management tools
# ---------------------------------------------------------------------------


@mcp.tool()
def mark_email(entry_id: str, read: bool = True) -> str:
    """Mark an email as read or unread."""
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        item.UnRead = not read
        item.Save()
        return f"Marked {item.Subject!r} as {'read' if read else 'unread'}."


@mcp.tool()
def move_email(entry_id: str, target_folder: str) -> str:
    """Move an email to another folder.

    Args:
        entry_id: EntryID of the email to move.
        target_folder: Destination folder name or path.
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        dest = resolve_folder(ns, target_folder)
        moved = item.Move(dest)
        return (
            f"Moved {item.Subject!r} to {dest.Name!r}. "
            f"New entry_id: {moved.EntryID}"
        )


@mcp.tool()
def delete_email(entry_id: str) -> str:
    """Delete an email (moves it to Deleted Items)."""
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        subject = item.Subject
        item.Delete()
        return f"Deleted: {subject!r} (moved to Deleted Items)."


# ---------------------------------------------------------------------------
# Calendar tools
# ---------------------------------------------------------------------------


@mcp.tool()
def list_calendar_events(days_ahead: int = 7, days_back: int = 0) -> str:
    """List calendar events in a date window around today.

    Args:
        days_ahead: How many days into the future to include.
        days_back: How many days into the past to include.
    """
    start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(
        days=days_back
    )
    end = start + timedelta(days=days_back + days_ahead + 1)
    with outlook_session() as ns:
        items = ns.GetDefaultFolder(FOLDER_CALENDAR).Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        restriction = (
            f"[Start] >= '{restrict_date(start)}' AND [Start] < '{restrict_date(end)}'"
        )
        items = items.Restrict(restriction)
        lines = []
        for i, appt in enumerate(items, 1):
            if i > 100:
                lines.append("[... more than 100 events, list truncated ...]")
                break
            location = f" @ {appt.Location}" if appt.Location else ""
            lines.append(
                f"{i}. {appt.Subject}{location}\n"
                f"   {fmt_dt(appt.Start)} -> {fmt_dt(appt.End)}\n"
                f"   entry_id: {appt.EntryID}"
            )
        return "\n\n".join(lines) if lines else "No events in this window."


@mcp.tool()
def create_calendar_event(
    subject: str,
    start: str,
    duration_minutes: int = 30,
    location: str = "",
    body: str = "",
    attendees: str = "",
) -> str:
    """Create a calendar event, optionally sending invites to attendees.

    Args:
        subject: Event title.
        start: Start time as "YYYY-MM-DD HH:MM".
        duration_minutes: Length of the event in minutes.
        location: Optional location.
        body: Optional description.
        attendees: Optional semicolon-separated attendee addresses; if given,
            the event is sent as a meeting invitation.
    """
    start_dt = parse_when(start)
    with outlook_session() as ns:
        app = ns.Application
        appt = app.CreateItem(ITEM_APPOINTMENT)
        appt.Subject = subject
        appt.Start = start_dt
        appt.Duration = duration_minutes
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        if attendees:
            appt.MeetingStatus = 1  # olMeeting
            for addr in attendees.split(";"):
                addr = addr.strip()
                if addr:
                    appt.Recipients.Add(addr)
            appt.Recipients.ResolveAll()
            appt.Send()
            return f"Meeting invite sent: {subject!r} at {start} to {attendees}"
        appt.Save()
        return f"Event created: {subject!r} at {start} ({duration_minutes} min)"


# ---------------------------------------------------------------------------
# Free/busy & scheduling (other people's availability)
# ---------------------------------------------------------------------------


@mcp.tool()
def check_availability(
    attendees: str,
    date: str,
    day_start_hour: int = 8,
    day_end_hour: int = 18,
    granularity_minutes: int = 30,
) -> str:
    """Show each attendee's free/busy timeline for a single day.

    Reads published free/busy data from Exchange/Outlook — no special mailbox
    permissions needed beyond what your organization already shares (typically
    everyone can see free/busy blocks, though not the meeting subjects).

    Args:
        attendees: Semicolon-separated names or email addresses to check.
            You can include yourself. Example: "alice@corp.com; Bob Smith".
        date: The day to inspect, as "YYYY-MM-DD".
        day_start_hour: First hour of the working window to display (0-23).
        day_end_hour: Last hour of the working window to display (1-24).
        granularity_minutes: Slot size; 15, 30, or 60 are typical.
    """
    if granularity_minutes < 1 or 1440 % granularity_minutes != 0:
        raise ValueError("granularity_minutes must divide 1440 (e.g. 15, 30, 60).")
    day = parse_date_only(date)
    people = [a.strip() for a in attendees.split(";") if a.strip()]
    if not people:
        raise ValueError("Provide at least one attendee.")

    start_idx = day_start_hour * 60 // granularity_minutes
    end_idx = day_end_hour * 60 // granularity_minutes

    with outlook_session() as ns:
        blocks = []
        for person in people:
            recip = resolve_recipient(ns, person)
            slots = freebusy_slots(recip, day, granularity_minutes)
            lines = [f"{recip.Name} <{recip.Address}> — {date}"]
            busy_ranges: list[tuple[datetime, datetime]] = []
            for idx in range(start_idx, min(end_idx, len(slots))):
                code = slots[idx]
                if code != FREEBUSY_FREE:
                    busy_ranges.append(
                        (
                            slot_time(day, idx, granularity_minutes),
                            slot_time(day, idx + 1, granularity_minutes),
                        )
                    )
            if busy_ranges:
                merged = merge_intervals(busy_ranges)
                lines.append("  Busy:")
                for bstart, bend in merged:
                    lines.append(
                        f"    {bstart.strftime('%H:%M')}–{bend.strftime('%H:%M')}"
                    )
            else:
                lines.append(
                    f"  Free all day ({day_start_hour:02d}:00–{day_end_hour:02d}:00)."
                )
            blocks.append("\n".join(lines))
        return "\n\n".join(blocks)


@mcp.tool()
def find_meeting_times(
    attendees: str,
    date: str,
    duration_minutes: int = 30,
    day_start_hour: int = 9,
    day_end_hour: int = 17,
    granularity_minutes: int = 30,
    max_suggestions: int = 10,
) -> str:
    """Find time slots on a given day when ALL attendees are free.

    Intersects everyone's free/busy data and returns open windows long enough
    for the requested meeting duration. Use this to answer "when can we all
    meet?" before calling create_calendar_event.

    Args:
        attendees: Semicolon-separated names or emails. Include everyone who
            must attend (add yourself too).
        date: The day to search, as "YYYY-MM-DD".
        duration_minutes: Required meeting length.
        day_start_hour: Earliest hour to consider (0-23).
        day_end_hour: Latest hour to consider (1-24).
        granularity_minutes: Slot resolution; must divide the duration cleanly
            for best results (15 or 30 recommended).
        max_suggestions: Cap on how many candidate start times to return.
    """
    if granularity_minutes < 1 or 1440 % granularity_minutes != 0:
        raise ValueError("granularity_minutes must divide 1440 (e.g. 15, 30, 60).")
    day = parse_date_only(date)
    people = [a.strip() for a in attendees.split(";") if a.strip()]
    if not people:
        raise ValueError("Provide at least one attendee.")

    start_idx = day_start_hour * 60 // granularity_minutes
    end_idx = day_end_hour * 60 // granularity_minutes
    needed = duration_minutes // granularity_minutes
    if duration_minutes % granularity_minutes != 0:
        needed += 1

    with outlook_session() as ns:
        resolved_names = []
        # combined[idx] is True when EVERY attendee is free in that slot.
        combined = [True] * end_idx
        unknown = set()
        for person in people:
            recip = resolve_recipient(ns, person)
            resolved_names.append(recip.Name)
            slots = freebusy_slots(recip, day, granularity_minutes)
            for idx in range(start_idx, end_idx):
                if idx >= len(slots):
                    # No published data this far out; treat as unknown/busy.
                    combined[idx] = False
                    unknown.add(recip.Name)
                elif slots[idx] != FREEBUSY_FREE:
                    combined[idx] = False

        suggestions: list[str] = []
        idx = start_idx
        while idx <= end_idx - needed and len(suggestions) < max_suggestions:
            if all(combined[idx : idx + needed]):
                s = slot_time(day, idx, granularity_minutes)
                e = slot_time(day, idx + needed, granularity_minutes)
                suggestions.append(f"  {s.strftime('%H:%M')}–{e.strftime('%H:%M')}")
                idx += needed  # jump past this slot to avoid overlapping picks
            else:
                idx += 1

        header = (
            f"Common {duration_minutes}-minute openings for "
            f"{', '.join(resolved_names)} on {date} "
            f"({day_start_hour:02d}:00–{day_end_hour:02d}:00):"
        )
        if not suggestions:
            body = "  None found — everyone is busy or free/busy data is unavailable."
        else:
            body = "\n".join(suggestions)
        note = ""
        if unknown:
            note = (
                "\n\nNote: no published free/busy data for "
                f"{', '.join(sorted(unknown))} in this window; "
                "those times were treated as unavailable."
            )
        return f"{header}\n{body}{note}"


@mcp.tool()
def schedule_out_of_office(
    start: str,
    end: str,
    subject: str = "Out of Office",
    all_day: bool = True,
) -> str:
    """Block a period on your calendar with 'Out of Office' availability status.

    This sets how you appear to others' free/busy (e.g. in check_availability /
    find_meeting_times) for the given dates. It marks you Out of Office and does
    NOT send meeting invitations to anyone.

    Note: this schedules your *status*, which is different from Outlook's
    automatic-reply (auto-responder) emails — see set_automatic_replies for why
    those cannot be configured through COM.

    Args:
        start: First day/time out, "YYYY-MM-DD" or "YYYY-MM-DD HH:MM".
        end: Return day/time. For all-day blocks this is the last full day out.
        subject: Title of the calendar block.
        all_day: If True, create an all-day block spanning start..end inclusive.
    """
    start_dt = parse_when(start)
    end_dt = parse_when(end)
    with outlook_session() as ns:
        app = ns.Application
        appt = app.CreateItem(ITEM_APPOINTMENT)
        appt.Subject = subject
        appt.BusyStatus = BUSY_OUT_OF_OFFICE
        appt.ReminderSet = False
        if all_day:
            appt.AllDayEvent = True
            appt.Start = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            # Outlook's all-day End is exclusive, so add a day past the last day out.
            appt.End = end_dt.replace(
                hour=0, minute=0, second=0, microsecond=0
            ) + timedelta(days=1)
        else:
            appt.Start = start_dt
            appt.End = end_dt
        appt.Save()
        span = f"{start} through {end}" if all_day else f"{start} to {end}"
        return f"Out-of-Office status block created: {subject!r}, {span}."


@mcp.tool()
def set_automatic_replies(
    message: str = "",
    start: str = "",
    end: str = "",
) -> str:
    """Explain how to configure Outlook automatic replies (auto-responder).

    IMPORTANT: Outlook's automatic replies / Out-of-Office Assistant cannot be
    set through COM automation — the Outlook Object Model does not expose OOF
    reply settings. This tool does not change any setting; it returns guidance.

    Configuring auto-replies requires one of:
      * Exchange Web Services (EWS) SetUserOofSettings, or
      * Microsoft Graph (mailboxSettings.automaticRepliesSetting),
    neither of which is part of win32com/COM. Use schedule_out_of_office to set
    your availability status, and set the auto-reply text manually in Outlook
    (File > Automatic Replies) or via one of the APIs above.
    """
    return (
        "Automatic replies cannot be set via COM / win32com — the Outlook "
        "Object Model does not expose Out-of-Office reply settings.\n\n"
        "To turn on auto-replies, either:\n"
        "  1. In classic Outlook: File > Automatic Replies (Out of Office), "
        "set the date range and message, then OK; or\n"
        "  2. Use Exchange Web Services (SetUserOofSettings) or Microsoft "
        "Graph (mailboxSettings.automaticRepliesSetting) — both outside COM.\n\n"
        "What this server CAN do without those APIs: schedule_out_of_office "
        "blocks your calendar with Out-of-Office availability status so you "
        "show as away in others' free/busy."
        + (
            f"\n\n(Requested message/dates were not applied — message={message!r}, "
            f"start={start!r}, end={end!r}.)"
            if (message or start or end)
            else ""
        )
    )


# ---------------------------------------------------------------------------
# Contacts
# ---------------------------------------------------------------------------


@mcp.tool()
def list_contacts(search: str = "", count: int = 50) -> str:
    """List contacts from the default Contacts folder, optionally filtered.

    Args:
        search: Optional substring to match against name/email/company.
        count: Maximum contacts to return (1-200).
    """
    count = max(1, min(count, 200))
    needle = search.lower()
    with outlook_session() as ns:
        items = ns.GetDefaultFolder(FOLDER_CONTACTS).Items
        lines = []
        for contact in items:
            if getattr(contact, "Class", None) != 40:  # olContact
                continue
            name = contact.FullName or ""
            email = contact.Email1Address or ""
            company = contact.CompanyName or ""
            if needle and needle not in f"{name} {email} {company}".lower():
                continue
            entry = name
            if email:
                entry += f" <{email}>"
            if company:
                entry += f" ({company})"
            lines.append(f"- {entry}")
            if len(lines) >= count:
                break
        return "\n".join(lines) if lines else "No contacts found."


def main() -> None:
    if sys.platform != "win32":
        print(
            "warning: outlook-mcp uses win32com and only works on Windows "
            "with classic Outlook desktop installed.",
            file=sys.stderr,
        )
    mcp.run()


if __name__ == "__main__":
    main()
