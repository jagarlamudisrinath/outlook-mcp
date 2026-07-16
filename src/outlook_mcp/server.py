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

import hashlib
import mimetypes
import re
import sys
import threading
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterator

from mcp.server.fastmcp import FastMCP

from . import timezones as tz

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

# OlMeetingStatus / OlMeetingResponse
MEETING_MEETING = 1  # olMeeting
MEETING_REQUEST_CLASS = 53  # olMeetingRequest (a MeetingItem)
RESPONSE_TENTATIVE = 2  # olMeetingTentative
RESPONSE_ACCEPTED = 3  # olMeetingAccepted
RESPONSE_DECLINED = 4  # olMeetingDeclined

RESPONSE_MAP = {
    "accept": RESPONSE_ACCEPTED,
    "accepted": RESPONSE_ACCEPTED,
    "tentative": RESPONSE_TENTATIVE,
    "maybe": RESPONSE_TENTATIVE,
    "decline": RESPONSE_DECLINED,
    "declined": RESPONSE_DECLINED,
}

# OlRecurrenceType
RECUR_DAILY = 0
RECUR_WEEKLY = 1
RECUR_MONTHLY = 2
RECUR_YEARLY = 5

RECUR_TYPE_MAP = {
    "daily": RECUR_DAILY,
    "weekly": RECUR_WEEKLY,
    "monthly": RECUR_MONTHLY,
    "yearly": RECUR_YEARLY,
}

# OlDaysOfWeek bit flags
DAY_MASK = {
    "sun": 1, "sunday": 1,
    "mon": 2, "monday": 2,
    "tue": 4, "tues": 4, "tuesday": 4,
    "wed": 8, "wednesday": 8,
    "thu": 16, "thur": 16, "thurs": 16, "thursday": 16,
    "fri": 32, "friday": 32,
    "sat": 64, "saturday": 64,
}

# OlMeetingStatus
MEETING_NONMEETING = 0
MEETING_CANCELED = 5

# OlResponseStatus
RESPONSE_STATUS = {
    0: "None",
    1: "Organizer",
    2: "Tentative",
    3: "Accepted",
    4: "Declined",
    5: "Not yet responded",
}

# OlMeetingRecipientType
MEETING_RECIP_TYPE = {0: "Organizer", 1: "Required", 2: "Optional", 3: "Resource"}

# OlImportance / OlSensitivity
IMPORTANCE = {0: "Low", 1: "Normal", 2: "High"}
IMPORTANCE_IN = {"low": 0, "normal": 1, "high": 2}
SENSITIVITY = {0: "Normal", 1: "Personal", 2: "Private", 3: "Confidential"}
SENSITIVITY_IN = {"normal": 0, "personal": 1, "private": 2, "confidential": 3}

APPOINTMENT_ITEM_CLASS = 26  # olAppointment

# MAPI property tags via PropertyAccessor
PROP_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PROP_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

# Teams / online-meeting join links found in a meeting body.
_MEETING_LINK_RE = re.compile(
    r"https://teams\.microsoft\.com/l/meetup-join/[^\s>\"'\]]+"
    r"|https://[\w.-]*\.zoom\.us/j/[^\s>\"'\]]+"
    r"|https://[\w.-]*\.webex\.com/[^\s>\"'\]]+",
    re.IGNORECASE,
)

_UNSAFE_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def extract_meeting_links(text: str) -> list[str]:
    """Return de-duplicated Teams/Zoom/Webex join links found in text."""
    seen: dict[str, None] = {}
    for match in _MEETING_LINK_RE.findall(text or ""):
        seen.setdefault(match.rstrip(".,);"), None)
    return list(seen)


def sanitize_filename(name: str) -> str:
    """Make an attachment filename safe: strip paths and illegal characters."""
    base = (name or "").replace("\\", "/").split("/")[-1]
    base = _UNSAFE_FILENAME_RE.sub("_", base).strip().rstrip(".")
    return base or "unnamed"


def unique_path(directory: Path, filename: str) -> Path:
    """Return a path in `directory` for `filename`, avoiding collisions."""
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    n = 1
    while True:
        alt = directory / f"{stem} ({n}){suffix}"
        if not alt.exists():
            return alt
        n += 1


def parse_day_mask(days: str) -> int:
    """Turn a string like 'Mon,Wed,Fri' into an OlDaysOfWeek bitmask."""
    mask = 0
    for token in days.replace(";", ",").split(","):
        token = token.strip().lower()
        if not token:
            continue
        if token not in DAY_MASK:
            raise ValueError(f"Unknown weekday {token!r}. Use Mon..Sun.")
        mask |= DAY_MASK[token]
    return mask


def apply_recurrence(
    appt: Any,
    recurrence: str,
    interval: int,
    days: str,
    count: int,
    until: str,
    start_dt: datetime,
    duration_minutes: int,
) -> str:
    """Configure an appointment's RecurrencePattern. Returns a human summary.

    The appointment's Start/Duration must already be set; do not set Start after
    calling this (that resets the pattern in the Outlook Object Model).
    """
    rtype = recurrence.strip().lower()
    if rtype not in RECUR_TYPE_MAP:
        raise ValueError('recurrence must be "daily", "weekly", "monthly", or "yearly".')
    interval = max(1, interval)

    rp = appt.GetRecurrencePattern()
    rp.RecurrenceType = RECUR_TYPE_MAP[rtype]
    rp.Interval = interval

    detail = f"every {interval} {rtype.rstrip('ly')}(s)" if interval > 1 else rtype
    if rtype == "weekly":
        mask = parse_day_mask(days) if days else 0
        if mask:
            rp.DayOfWeekMask = mask
            detail += f" on {days}"
    elif rtype == "monthly":
        rp.DayOfMonth = start_dt.day
        detail += f" on day {start_dt.day}"

    rp.PatternStartDate = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
    if count and count > 0:
        rp.Occurrences = count
        detail += f", {count} occurrences"
    elif until:
        end_dt = parse_date_only(until)
        rp.PatternEndDate = end_dt
        detail += f", until {until}"
    else:
        detail += ", no end date"
    return detail

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


def split_recipients(value: str) -> list[str]:
    """Split a recipient string on semicolons/newlines into trimmed entries."""
    parts: list[str] = []
    for chunk in value.replace("\n", ";").split(";"):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    return parts


def recipient_smtp(recipient: Any) -> str:
    """Canonical SMTP address of a resolved Recipient (resolves Exchange EX)."""
    try:
        entry = recipient.AddressEntry
        if entry is not None and entry.Type == "EX":
            exch = entry.GetExchangeUser()
            if exch is not None and exch.PrimarySmtpAddress:
                return exch.PrimarySmtpAddress
        if entry is not None and entry.Address:
            return entry.Address
    except Exception:
        pass
    return getattr(recipient, "Address", "") or ""


def add_and_resolve_recipients(
    item: Any, buckets: list[tuple[str, str, int]], allow_unresolved: bool
) -> list[dict[str, str]]:
    """Add recipients to a mail item by type, resolve, and validate.

    buckets is a list of (address_string, label, olMailRecipientType). Raises
    ValueError listing any names that could not be resolved unless
    allow_unresolved is True. Returns one dict per recipient with its resolved
    name and canonical SMTP address.
    """
    added: list[tuple[str, str, Any]] = []
    for addr_string, label, rtype in buckets:
        for addr in split_recipients(addr_string):
            recip = item.Recipients.Add(addr)
            recip.Type = rtype
            added.append((label, addr, recip))

    item.Recipients.ResolveAll()

    unresolved = [addr for _, addr, recip in added if not recip.Resolved]
    if unresolved and not allow_unresolved:
        raise ValueError(
            "Could not resolve these recipients in the address book / GAL: "
            + ", ".join(repr(a) for a in unresolved)
            + ". Fix the addresses, or pass allow_unresolved=True to send anyway."
        )

    resolved: list[dict[str, str]] = []
    for label, addr, recip in added:
        resolved.append(
            {
                "type": label,
                "input": addr,
                "name": getattr(recip, "Name", addr) or addr,
                "smtp": recipient_smtp(recip) if recip.Resolved else "",
                "resolved": "yes" if recip.Resolved else "NO",
            }
        )
    return resolved


def format_recipient_lines(resolved: list[dict[str, str]]) -> str:
    lines = []
    for r in resolved:
        smtp = f" <{r['smtp']}>" if r["smtp"] else ""
        unresolved = "" if r["resolved"] == "yes" else "  [UNRESOLVED]"
        lines.append(f"  {r['type']}: {r['name']}{smtp}{unresolved}")
    return "\n".join(lines)


def fmt_dt(value: Any) -> str:
    try:
        return value.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(value)


def mail_summary(item: Any, folder: str = "") -> dict[str, Any]:
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "from": f"{item.SenderName} <{sender_address(item)}>",
        "to": item.To or "",
        "received": fmt_dt(item.ReceivedTime),
        "unread": bool(item.UnRead),
        "has_attachments": item.Attachments.Count > 0,
        "preview": (item.Body or "")[:200].replace("\r\n", " ").strip(),
        "folder": folder,
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
        folder = f"  |  Folder: {m['folder']}" if m.get("folder") else ""
        lines.append(
            f"{i}. {m['subject']}{flag_str}\n"
            f"   From: {m['from']}  |  Received: {m['received']}{folder}\n"
            f"   Preview: {m['preview']}\n"
            f"   entry_id: {m['entry_id']}"
        )
    return "\n\n".join(lines)


def prop_value(item: Any, schema: str) -> str:
    """Read a MAPI property via PropertyAccessor, returning '' on failure."""
    try:
        return item.PropertyAccessor.GetProperty(schema) or ""
    except Exception:
        return ""


def meeting_attendees(appt: Any) -> list[dict[str, str]]:
    """Return per-recipient role, name, SMTP and response status for a meeting."""
    people: list[dict[str, str]] = []
    try:
        for recip in appt.Recipients:
            try:
                rtype = MEETING_RECIP_TYPE.get(recip.Type, str(recip.Type))
                status = RESPONSE_STATUS.get(
                    getattr(recip, "MeetingResponseStatus", 0), "Unknown"
                )
            except Exception:
                rtype, status = "Unknown", "Unknown"
            people.append(
                {
                    "role": rtype,
                    "name": getattr(recip, "Name", "") or "",
                    "smtp": recipient_smtp(recip),
                    "response": status,
                }
            )
    except Exception:
        pass
    return people


def appointment_details(appt: Any) -> str:
    """Render a rich, human-readable description of a calendar item."""
    lines = [f"Subject: {appt.Subject}"]
    try:
        lines.append(f"Organizer: {appt.Organizer}")
    except Exception:
        pass
    lines.append(f"When: {fmt_dt(appt.Start)} -> {fmt_dt(appt.End)}")
    try:
        if appt.AllDayEvent:
            lines.append("All-day event")
    except Exception:
        pass
    if getattr(appt, "Location", ""):
        lines.append(f"Location: {appt.Location}")
    try:
        lines.append(f"Duration: {appt.Duration} min")
    except Exception:
        pass

    # Recurrence
    try:
        if appt.IsRecurring:
            rp = appt.GetRecurrencePattern()
            lines.append(f"Recurring: yes (pattern from {fmt_dt(rp.PatternStartDate)})")
    except Exception:
        pass

    # Status / classification
    try:
        lines.append(f"My response: {RESPONSE_STATUS.get(appt.ResponseStatus, '?')}")
    except Exception:
        pass
    for label, prop, table in (
        ("Importance", "Importance", IMPORTANCE),
        ("Sensitivity", "Sensitivity", SENSITIVITY),
    ):
        try:
            lines.append(f"{label}: {table.get(getattr(appt, prop), '?')}")
        except Exception:
            pass
    if getattr(appt, "Categories", ""):
        lines.append(f"Categories: {appt.Categories}")
    try:
        lines.append(f"Last modified: {fmt_dt(appt.LastModificationTime)}")
    except Exception:
        pass

    # Teams / online meeting links
    links = extract_meeting_links(getattr(appt, "Body", "") or "")
    if links:
        lines.append("Join links:")
        lines.extend(f"  {u}" for u in links)

    # Attendees
    attendees = meeting_attendees(appt)
    if attendees:
        lines.append("Attendees:")
        for a in attendees:
            smtp = f" <{a['smtp']}>" if a["smtp"] else ""
            lines.append(f"  [{a['role']}] {a['name']}{smtp} — {a['response']}")

    # Attachments
    try:
        if appt.Attachments.Count:
            names = [att.FileName for att in appt.Attachments]
            lines.append("Attachments: " + ", ".join(names))
    except Exception:
        pass

    lines.append(f"entry_id: {appt.EntryID}")
    body = (getattr(appt, "Body", "") or "").strip()
    if body:
        lines.append("")
        lines.append(body[:2000])
    return "\n".join(lines)


def dasl_escape(text: str) -> str:
    return text.replace("'", "''").replace("%", "[%]")


def parse_when(value: str) -> datetime:
    """Parse a naive local datetime (shared with the timezone module)."""
    return tz.parse_naive(value)


def apply_timezone(app: Any, appt: Any, start_dt: datetime, timezone_name: str) -> dict[str, str]:
    """Set an appointment's Start/End time zone via COM and report the times.

    Must be called before setting appt.Start (the COM object interprets Start as
    a wall-clock time in StartTimeZone). Returns a dict describing local/UTC
    times so callers can surface unambiguous, DST-correct results.
    """
    rtz = tz.resolve_timezone(timezone_name)
    aware = tz.localize(start_dt, rtz)
    utc = tz.to_utc(aware)
    info = {
        "timezone": rtz.label,
        "start_local": tz.fmt(aware),
        "start_utc": tz.fmt(utc),
    }
    if rtz.windows is not None:
        try:
            com_tz = app.TimeZones.Item(rtz.windows)
            appt.StartTimeZone = com_tz
            appt.EndTimeZone = com_tz
        except Exception as exc:  # pragma: no cover - COM-only path
            info["timezone_warning"] = (
                f"Could not apply Windows time zone {rtz.windows!r} in Outlook "
                f"({exc}); the time was set in the profile's zone instead."
            )
    else:
        info["timezone_warning"] = (
            f"No Windows time-zone ID mapped for {rtz.iana!r}; Outlook stored the "
            "time in the profile's zone. UTC value above is still correct."
        )
    return info


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


def _email_recipient_smtps(item: Any) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    try:
        for r in item.Recipients:
            out.append(
                {
                    "type": {1: "To", 2: "CC", 3: "BCC"}.get(
                        getattr(r, "Type", 0), "?"
                    ),
                    "name": getattr(r, "Name", "") or "",
                    "smtp": recipient_smtp(r),
                }
            )
    except Exception:
        pass
    return out


@mcp.tool()
def get_email(
    entry_id: str,
    body_max_chars: int = 8000,
    include_html: bool = False,
    include_headers: bool = False,
) -> str:
    """Read a full email by its entry_id, with rich metadata.

    Includes canonical SMTP addresses for every recipient, conversation ID,
    Internet Message-ID, attachment list with indexes, and any Teams/Zoom/Webex
    links found in the body. Optionally returns the HTML body and raw internet
    headers.

    Args:
        entry_id: The Outlook EntryID of the email.
        body_max_chars: Truncate the plain-text body after this many characters
            (set higher for long messages; 0 means no body).
        include_html: Also include the raw HTML body.
        include_headers: Also include the raw internet (transport) headers.
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        body = (item.Body or "").strip()
        truncated = len(body) > body_max_chars
        parts = [
            f"Subject: {item.Subject}",
            f"From: {item.SenderName} <{sender_address(item)}>",
        ]
        for r in _email_recipient_smtps(item):
            smtp = f" <{r['smtp']}>" if r["smtp"] else ""
            parts.append(f"{r['type']}: {r['name']}{smtp}")
        parts.append(f"Received: {fmt_dt(item.ReceivedTime)}")
        try:
            parts.append(f"Importance: {IMPORTANCE.get(item.Importance, '?')}")
        except Exception:
            pass
        try:
            if item.Categories:
                parts.append(f"Categories: {item.Categories}")
        except Exception:
            pass
        try:
            parts.append(f"Unread: {bool(item.UnRead)}")
        except Exception:
            pass
        try:
            parts.append(f"Conversation ID: {item.ConversationID}")
        except Exception:
            pass
        msg_id = prop_value(item, PROP_INTERNET_MESSAGE_ID)
        if msg_id:
            parts.append(f"Message-ID: {msg_id}")
        parts.append(f"entry_id: {item.EntryID}")

        try:
            if item.Attachments.Count:
                parts.append("Attachments:")
                for idx, a in enumerate(item.Attachments, 1):
                    parts.append(f"  [{idx}] {a.FileName} ({a.Size} bytes)")
        except Exception:
            pass

        links = extract_meeting_links(item.Body or "")
        if links:
            parts.append("Links: " + ", ".join(links))

        if include_headers:
            headers = prop_value(item, PROP_TRANSPORT_HEADERS)
            if headers:
                parts.append("\n--- Internet headers ---")
                parts.append(headers[:4000])

        if body_max_chars > 0:
            parts.append("")
            parts.append(body[:body_max_chars])
            if truncated:
                parts.append(f"\n[... body truncated at {body_max_chars} chars ...]")

        if include_html:
            try:
                parts.append("\n--- HTML body ---")
                parts.append((item.HTMLBody or "")[: max(body_max_chars, 2000)])
            except Exception:
                pass
        return "\n".join(parts)


def iter_mail_folders(root: Any, recurse: bool) -> Iterator[Any]:
    """Yield a folder and, if recurse, all of its subfolders."""
    yield root
    if recurse:
        try:
            for sub in root.Folders:
                yield from iter_mail_folders(sub, True)
        except Exception:
            pass


def build_email_dasl(
    query: str,
    search_in: str,
    sender: str,
    to: str,
    since: str,
    until: str,
    unread_only: bool,
    flagged_only: bool,
    has_attachments: bool,
    importance: str,
) -> str:
    """Assemble an @SQL DASL restriction from the given filters."""
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
    clauses: list[str] = []
    if query:
        q = dasl_escape(query)
        clauses.append(
            "(" + " OR ".join(f'"{f}" LIKE \'%{q}%\'' for f in fields[search_in]) + ")"
        )
    if sender:
        s = dasl_escape(sender)
        clauses.append(
            '("urn:schemas:httpmail:fromemail" LIKE \'%' + s + "%' OR "
            '"urn:schemas:httpmail:fromname" LIKE \'%' + s + "%')"
        )
    if to:
        t = dasl_escape(to)
        clauses.append(
            '("urn:schemas:httpmail:to" LIKE \'%' + t + "%' OR "
            '"urn:schemas:httpmail:cc" LIKE \'%' + t + "%')"
        )
    if since:
        clauses.append(
            f'"urn:schemas:httpmail:datereceived" >= '
            f"'{parse_when(since).strftime('%Y/%m/%d %H:%M')}'"
        )
    if until:
        clauses.append(
            f'"urn:schemas:httpmail:datereceived" <= '
            f"'{parse_when(until).strftime('%Y/%m/%d %H:%M')}'"
        )
    if unread_only:
        clauses.append('"urn:schemas:httpmail:read" = 0')
    if has_attachments:
        clauses.append('"urn:schemas:httpmail:hasattachment" = 1')
    if flagged_only:
        clauses.append(
            '"http://schemas.microsoft.com/mapi/proptag/0x10900003" = 2'
        )
    if importance:
        if importance.lower() not in IMPORTANCE_IN:
            raise ValueError("importance must be low, normal, or high.")
        clauses.append(
            f'"urn:schemas:httpmail:importance" = {IMPORTANCE_IN[importance.lower()]}'
        )
    if not clauses:
        raise ValueError("Provide at least one search filter.")
    return "@SQL=(" + " AND ".join(clauses) + ")"


@mcp.tool()
def search_emails(
    query: str = "",
    folder: str = "Inbox",
    count: int = 20,
    search_in: str = "all",
    sender: str = "",
    to: str = "",
    since: str = "",
    until: str = "",
    unread_only: bool = False,
    flagged_only: bool = False,
    has_attachments: bool = False,
    importance: str = "",
    all_folders: bool = False,
    include_subfolders: bool = False,
) -> str:
    """Search emails with rich filters, in one folder or across all mailboxes.

    Args:
        query: Text to search for (matched per search_in). May be empty if other
            filters are given.
        folder: Folder to search (ignored when all_folders=True).
        count: Maximum results (1-200).
        search_in: Where to match `query`: "subject", "from", "body", or "all".
        sender: Only mail whose sender name/address contains this.
        to: Only mail whose To/CC contains this.
        since: Only mail received on/after this "YYYY-MM-DD [HH:MM]".
        until: Only mail received on/before this "YYYY-MM-DD [HH:MM]".
        unread_only: Only unread mail.
        flagged_only: Only flagged mail.
        has_attachments: Only mail with attachments.
        importance: "low", "normal", or "high".
        all_folders: Search every folder in every account/store.
        include_subfolders: When searching a single folder, also search beneath it.
    """
    count = max(1, min(count, 200))
    dasl = build_email_dasl(
        query, search_in, sender, to, since, until,
        unread_only, flagged_only, has_attachments, importance,
    )
    with outlook_session() as ns:
        if all_folders:
            roots = []
            for store_folder in ns.Folders:
                roots.append((store_folder, True))
        else:
            roots = [(resolve_folder(ns, folder), include_subfolders)]

        results: list[dict[str, Any]] = []
        for root, recurse in roots:
            for fld in iter_mail_folders(root, recurse):
                try:
                    items = fld.Items
                    items.Sort("[ReceivedTime]", True)
                    items = items.Restrict(dasl)
                except Exception:
                    continue
                for item in items:
                    if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                        continue
                    results.append(mail_summary(item, folder=fld.Name))
                    if len(results) >= count:
                        break
                if len(results) >= count:
                    break
            if len(results) >= count:
                break
        return render_mail_list(results)


@mcp.tool()
def list_attachments(entry_id: str) -> str:
    """List an email's attachments with index, size, and inline/real status.

    Args:
        entry_id: The email's EntryID.
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        if item.Attachments.Count == 0:
            return "This email has no attachments."
        lines = []
        for idx, att in enumerate(item.Attachments, 1):
            # Type 6 = olEmbeddedItem; inline images usually have a content id.
            cid = prop_value(
                att, "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
            )
            kind = "inline" if cid else "attachment"
            mime, _ = mimetypes.guess_type(att.FileName)
            lines.append(
                f"[{idx}] {att.FileName} — {att.Size} bytes, {kind}"
                f"{', ' + mime if mime else ''}"
            )
        return "\n".join(lines)


@mcp.tool()
def save_attachments(
    entry_id: str,
    save_dir: str,
    indexes: list[int] | None = None,
    max_size_mb: int = 50,
    include_inline: bool = False,
    overwrite: bool = False,
) -> str:
    """Save an email's attachments to a directory, safely.

    Filenames are sanitized (path components and illegal characters stripped),
    collisions get a numbered suffix unless overwrite=True, oversized files are
    skipped, and a SHA-256 hash is reported for each saved file.

    Args:
        entry_id: The email's EntryID.
        save_dir: Directory to save into (created if missing).
        indexes: 1-based attachment indexes to save (from list_attachments);
            omit to save all.
        max_size_mb: Skip attachments larger than this many megabytes.
        include_inline: Also save inline images (skipped by default).
        overwrite: Overwrite existing files instead of adding a numbered suffix.
    """
    target = Path(save_dir).expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)
    limit = max_size_mb * 1024 * 1024
    wanted = set(indexes) if indexes else None
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        if item.Attachments.Count == 0:
            return "This email has no attachments."
        saved, skipped = [], []
        for idx, att in enumerate(item.Attachments, 1):
            if wanted is not None and idx not in wanted:
                continue
            cid = prop_value(
                att, "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
            )
            if cid and not include_inline:
                skipped.append(f"[{idx}] {att.FileName} (inline image)")
                continue
            if att.Size > limit:
                skipped.append(
                    f"[{idx}] {att.FileName} ({att.Size} bytes > {max_size_mb} MB)"
                )
                continue
            safe = sanitize_filename(att.FileName)
            dest = target / safe if overwrite else unique_path(target, safe)
            att.SaveAsFile(str(dest))
            digest = hashlib.sha256(dest.read_bytes()).hexdigest()[:16]
            mime, _ = mimetypes.guess_type(dest.name)
            saved.append(f"{dest}  (sha256:{digest}{', ' + mime if mime else ''})")

        out = []
        if saved:
            out.append("Saved:\n" + "\n".join(f"  {s}" for s in saved))
        if skipped:
            out.append("Skipped:\n" + "\n".join(f"  {s}" for s in skipped))
        return "\n\n".join(out) if out else "Nothing saved."


# ---------------------------------------------------------------------------
# Email writing tools
# ---------------------------------------------------------------------------


def find_account(app: Any, identifier: str) -> Any:
    """Find an Outlook Account by SMTP address or display name."""
    ident = identifier.strip().lower()
    accounts = app.Session.Accounts
    for i in range(1, accounts.Count + 1):
        acct = accounts.Item(i)
        smtp = (getattr(acct, "SmtpAddress", "") or "").lower()
        name = (getattr(acct, "DisplayName", "") or "").lower()
        if ident in (smtp, name):
            return acct
    raise ValueError(
        f"No Outlook account matches {identifier!r}. Use list_accounts to see "
        "the available sending accounts."
    )


def apply_mail_options(
    app: Any,
    mail: Any,
    importance: str,
    sensitivity: str,
    categories: str,
    account: str,
    send_on_behalf: str,
    schedule_send: str,
) -> list[str]:
    """Apply optional classification/account/scheduling fields. Returns notes."""
    notes: list[str] = []
    if importance:
        if importance.lower() not in IMPORTANCE_IN:
            raise ValueError("importance must be low, normal, or high.")
        mail.Importance = IMPORTANCE_IN[importance.lower()]
        notes.append(f"importance={importance}")
    if sensitivity:
        if sensitivity.lower() not in SENSITIVITY_IN:
            raise ValueError(
                "sensitivity must be normal, personal, private, or confidential."
            )
        mail.Sensitivity = SENSITIVITY_IN[sensitivity.lower()]
        notes.append(f"sensitivity={sensitivity}")
    if categories:
        mail.Categories = categories
        notes.append(f"categories={categories}")
    if account:
        mail.SendUsingAccount = find_account(app, account)
        notes.append(f"send-as {account}")
    if send_on_behalf:
        mail.SentOnBehalfOfName = send_on_behalf
        notes.append(f"on behalf of {send_on_behalf}")
    if schedule_send:
        mail.DeferredDeliveryTime = parse_when(schedule_send)
        notes.append(f"deferred until {schedule_send}")
    return notes


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
    allow_unresolved: bool = False,
    importance: str = "",
    sensitivity: str = "",
    categories: str = "",
    account: str = "",
    send_on_behalf: str = "",
    schedule_send: str = "",
    request_read_receipt: bool = False,
) -> str:
    """Send a new email (or save it as a draft) through Outlook.

    Every recipient is added individually and resolved against the address book
    before sending. If any name is ambiguous or unknown, the send is rejected
    with the offending entries listed (unless allow_unresolved=True). The result
    lists each recipient's canonical SMTP address so you can confirm before
    trusting the send.

    Args:
        to: Recipient address(es), separated by semicolons.
        subject: Email subject.
        body: Email body text (or HTML if html=True).
        cc: CC address(es), semicolon-separated.
        bcc: BCC address(es), semicolon-separated.
        html: Treat body as HTML.
        attachments: Absolute paths of files to attach.
        save_as_draft: Save to Drafts instead of sending.
        allow_unresolved: Send even if some recipients don't resolve.
        importance: "low", "normal", or "high".
        sensitivity: "normal", "personal", "private", or "confidential".
        categories: Comma-separated Outlook categories to tag the message.
        account: SMTP address or display name of the account to send AS
            (see list_accounts). Requires you have that account/permission.
        send_on_behalf: Address to send ON BEHALF OF (needs delegate rights).
        schedule_send: Defer delivery until "YYYY-MM-DD HH:MM" (Outlook must be
            running/connected at that time for Exchange to release it).
        request_read_receipt: Ask for a read receipt.
    """
    if not split_recipients(to):
        raise ValueError("At least one 'to' recipient is required.")
    with outlook_session() as ns:
        app = ns.Application
        mail = app.CreateItem(ITEM_MAIL)
        resolved = add_and_resolve_recipients(
            mail,
            [(to, "To", 1), (cc, "CC", 2), (bcc, "BCC", 3)],
            allow_unresolved,
        )
        mail.Subject = subject
        if html:
            mail.HTMLBody = body
        else:
            mail.Body = body
        if request_read_receipt:
            mail.ReadReceiptRequested = True
        notes = apply_mail_options(
            app, mail, importance, sensitivity, categories,
            account, send_on_behalf, schedule_send,
        )
        for path in attachments or []:
            file = Path(path).expanduser()
            if not file.is_file():
                raise ValueError(f"Attachment not found: {path}")
            mail.Attachments.Add(str(file))

        recipient_block = format_recipient_lines(resolved)
        opts = ("\nOptions: " + ", ".join(notes)) if notes else ""
        if save_as_draft:
            mail.Save()
            return (
                f"Draft saved: {subject!r}\nentry_id: {mail.EntryID}\n"
                f"Recipients:\n{recipient_block}{opts}"
            )
        mail.Send()
        verb = "Email scheduled" if schedule_send else "Email sent"
        return f"{verb}: {subject!r}\nRecipients:\n{recipient_block}{opts}"


@mcp.tool()
def reply_to_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    save_as_draft: bool = False,
    attachments: list[str] | None = None,
) -> str:
    """Reply to an email. The original message is quoted below your reply.

    Args:
        entry_id: EntryID of the email to reply to.
        body: Your reply text (HTML if html=True).
        reply_all: Reply to all recipients instead of just the sender.
        html: Treat body as HTML (prepended above the quoted original).
        save_as_draft: Save to Drafts instead of sending.
        attachments: Absolute paths of files to attach to the reply.
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        reply = item.ReplyAll() if reply_all else item.Reply()
        if html:
            reply.HTMLBody = body + "<br><br>" + (reply.HTMLBody or "")
        else:
            reply.Body = body + "\n\n" + reply.Body
        for path in attachments or []:
            file = Path(path).expanduser()
            if not file.is_file():
                raise ValueError(f"Attachment not found: {path}")
            reply.Attachments.Add(str(file))
        if save_as_draft:
            reply.Save()
            return f"Reply draft saved for: {item.Subject!r}\nentry_id: {reply.EntryID}"
        reply.Send()
        return f"Reply sent for: {item.Subject!r}"


@mcp.tool()
def forward_email(
    entry_id: str,
    to: str,
    cc: str = "",
    bcc: str = "",
    comment: str = "",
    html: bool = False,
    save_as_draft: bool = False,
    allow_unresolved: bool = False,
) -> str:
    """Forward an email (attachments included), with validated recipients.

    Args:
        entry_id: EntryID of the email to forward.
        to: Recipient address(es), semicolon-separated.
        cc: CC address(es), semicolon-separated.
        bcc: BCC address(es), semicolon-separated.
        comment: Optional text placed above the forwarded message.
        html: Treat comment as HTML.
        save_as_draft: Save to Drafts instead of sending.
        allow_unresolved: Forward even if some recipients don't resolve.
    """
    if not split_recipients(to):
        raise ValueError("At least one 'to' recipient is required.")
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        fwd = item.Forward()
        resolved = add_and_resolve_recipients(
            fwd, [(to, "To", 1), (cc, "CC", 2), (bcc, "BCC", 3)], allow_unresolved
        )
        if comment:
            if html:
                fwd.HTMLBody = comment + "<br><br>" + (fwd.HTMLBody or "")
            else:
                fwd.Body = comment + "\n\n" + fwd.Body
        block = format_recipient_lines(resolved)
        if save_as_draft:
            fwd.Save()
            return (
                f"Forward draft saved for: {item.Subject!r}\n"
                f"entry_id: {fwd.EntryID}\nRecipients:\n{block}"
            )
        fwd.Send()
        return f"Forwarded {item.Subject!r}\nRecipients:\n{block}"


# ---------------------------------------------------------------------------
# Draft lifecycle
# ---------------------------------------------------------------------------


@mcp.tool()
def list_drafts(count: int = 20) -> str:
    """List messages in the Drafts folder.

    Args:
        count: Maximum drafts to return (1-100).
    """
    count = max(1, min(count, 100))
    with outlook_session() as ns:
        items = resolve_folder(ns, "Drafts").Items
        results = []
        for item in items:
            if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                continue
            results.append(
                f"{len(results) + 1}. {item.Subject or '(no subject)'}\n"
                f"   To: {item.To or '(none)'}\n"
                f"   entry_id: {item.EntryID}"
            )
            if len(results) >= count:
                break
        return "\n\n".join(results) if results else "No drafts."


@mcp.tool()
def update_draft(
    entry_id: str,
    subject: str | None = None,
    body: str | None = None,
    html: bool = False,
    to: str | None = None,
    cc: str | None = None,
    bcc: str | None = None,
    add_attachments: list[str] | None = None,
) -> str:
    """Edit a saved draft in place (subject, body, recipients, attachments).

    Only the fields you pass are changed. Passing to/cc/bcc REPLACES that
    recipient set. Returns the draft's entry_id.

    Args:
        entry_id: EntryID of the draft.
        subject: New subject.
        body: New body (HTML if html=True).
        html: Treat body as HTML.
        to: Replace the To recipients (semicolon-separated).
        cc: Replace the CC recipients.
        bcc: Replace the BCC recipients.
        add_attachments: Absolute paths of files to attach.
    """
    with outlook_session() as ns:
        mail = get_item(ns, entry_id)
        changed = []
        if subject is not None:
            mail.Subject = subject
            changed.append("subject")
        if body is not None:
            if html:
                mail.HTMLBody = body
            else:
                mail.Body = body
            changed.append("body")
        if to is not None or cc is not None or bcc is not None:
            for i in range(mail.Recipients.Count, 0, -1):
                mail.Recipients.Remove(i)
            buckets = []
            if to:
                buckets.append((to, "To", 1))
            if cc:
                buckets.append((cc, "CC", 2))
            if bcc:
                buckets.append((bcc, "BCC", 3))
            if buckets:
                add_and_resolve_recipients(mail, buckets, allow_unresolved=False)
            changed.append("recipients")
        for path in add_attachments or []:
            file = Path(path).expanduser()
            if not file.is_file():
                raise ValueError(f"Attachment not found: {path}")
            mail.Attachments.Add(str(file))
            changed.append("attachment")
        mail.Save()
        return f"Draft updated ({', '.join(changed) or 'no changes'}): entry_id {mail.EntryID}"


@mcp.tool()
def send_draft(entry_id: str, allow_unresolved: bool = False) -> str:
    """Send an existing draft. Recipients are re-resolved and validated first.

    Args:
        entry_id: EntryID of the draft to send.
        allow_unresolved: Send even if some recipients don't resolve.
    """
    with outlook_session() as ns:
        mail = get_item(ns, entry_id)
        if not split_recipients(mail.To or ""):
            raise ValueError("Draft has no 'To' recipient.")
        mail.Recipients.ResolveAll()
        unresolved = [
            r.Name for r in mail.Recipients if not r.Resolved
        ]
        if unresolved and not allow_unresolved:
            raise ValueError(
                "Unresolved recipients: " + ", ".join(unresolved)
                + ". Fix them or pass allow_unresolved=True."
            )
        subject = mail.Subject
        mail.Send()
        return f"Draft sent: {subject!r}"


@mcp.tool()
def delete_draft(entry_id: str) -> str:
    """Delete a draft (moves it to Deleted Items)."""
    with outlook_session() as ns:
        mail = get_item(ns, entry_id)
        subject = mail.Subject
        mail.Delete()
        return f"Draft deleted: {subject!r}"


@mcp.tool()
def remove_attachment(entry_id: str, index: int) -> str:
    """Remove one attachment (1-based index) from a draft or item.

    Args:
        entry_id: EntryID of the mail item.
        index: 1-based attachment index (from list_attachments).
    """
    with outlook_session() as ns:
        item = get_item(ns, entry_id)
        if index < 1 or index > item.Attachments.Count:
            raise ValueError(
                f"index {index} out of range (item has {item.Attachments.Count})."
            )
        name = item.Attachments.Item(index).FileName
        item.Attachments.Remove(index)
        item.Save()
        return f"Removed attachment {name!r} from {item.Subject!r}."


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
    timezone_name: str = "",
) -> str:
    """Create a calendar event, optionally sending invites to attendees.

    Args:
        subject: Event title.
        start: Start time as "YYYY-MM-DD HH:MM" (wall-clock in timezone_name).
        duration_minutes: Length of the event in minutes.
        location: Optional location.
        body: Optional description.
        attendees: Optional semicolon-separated attendee addresses; if given,
            the event is sent as a meeting invitation.
        timezone_name: Time zone for `start`, as an IANA name
            ("Asia/Kolkata"), Windows ID ("India Standard Time"), or common
            abbreviation ("IST"). If omitted, the time is interpreted in
            Outlook's own profile time zone (may be ambiguous across DST).
    """
    start_dt = parse_when(start)
    with outlook_session() as ns:
        app = ns.Application
        appt = app.CreateItem(ITEM_APPOINTMENT)
        appt.Subject = subject
        tzinfo: dict[str, str] = {}
        if timezone_name:
            tzinfo = apply_timezone(app, appt, start_dt, timezone_name)
        appt.Start = start_dt
        appt.Duration = duration_minutes
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        tz_note = _tz_note(tzinfo)
        if attendees:
            appt.MeetingStatus = MEETING_MEETING
            resolved = add_and_resolve_recipients(
                appt, [(attendees, "Required", 1)], allow_unresolved=False
            )
            appt.Send()
            return (
                f"Meeting invite sent: {subject!r} at {start}{tz_note}\n"
                f"Attendees:\n{format_recipient_lines(resolved)}"
            )
        appt.Save()
        return f"Event created: {subject!r} at {start} ({duration_minutes} min){tz_note}"


def _tz_note(tzinfo: dict[str, str]) -> str:
    """Render an apply_timezone() result dict as a short human-readable suffix."""
    if not tzinfo:
        return ""
    note = (
        f"\nTime zone: {tzinfo['timezone']}"
        f"\nLocal:  {tzinfo['start_local']}"
        f"\nUTC:    {tzinfo['start_utc']}"
    )
    if "timezone_warning" in tzinfo:
        note += f"\nWarning: {tzinfo['timezone_warning']}"
    return note


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
def _free_slots_for_day(
    ns: Any, people: list[str], day: datetime, granularity_minutes: int,
    start_idx: int, end_idx: int,
) -> tuple[list[bool], set[str], list[str]]:
    """Return per-slot combined availability for one day plus unknown names."""
    combined = [True] * end_idx
    unknown: set[str] = set()
    names: list[str] = []
    for person in people:
        recip = resolve_recipient(ns, person)
        if recip.Name not in names:
            names.append(recip.Name)
        slots = freebusy_slots(recip, day, granularity_minutes)
        for idx in range(start_idx, end_idx):
            if idx >= len(slots):
                combined[idx] = False
                unknown.add(recip.Name)
            elif slots[idx] != FREEBUSY_FREE:
                combined[idx] = False
    return combined, unknown, names


@mcp.tool()
def find_meeting_times(
    attendees: str,
    date: str,
    duration_minutes: int = 30,
    days_to_search: int = 1,
    day_start_hour: int = 9,
    day_end_hour: int = 17,
    granularity_minutes: int = 30,
    max_suggestions: int = 10,
    include_me: bool = True,
) -> str:
    """Find time slots when ALL attendees are free, over one or more days.

    Intersects everyone's free/busy data and returns open windows long enough
    for the requested meeting duration. Use this to answer "when can we all
    meet?" before calling create_calendar_event.

    Args:
        attendees: Semicolon-separated names or emails of the required attendees.
        date: First day to search, as "YYYY-MM-DD".
        duration_minutes: Required meeting length.
        days_to_search: How many consecutive days from `date` to scan (1-31).
        day_start_hour: Earliest hour to consider each day (0-23).
        day_end_hour: Latest hour to consider each day (1-24).
        granularity_minutes: Slot resolution (15 or 30 recommended).
        max_suggestions: Cap on how many candidate start times to return.
        include_me: Also require the current user (you) to be free (default True).
    """
    if granularity_minutes < 1 or 1440 % granularity_minutes != 0:
        raise ValueError("granularity_minutes must divide 1440 (e.g. 15, 30, 60).")
    days_to_search = max(1, min(days_to_search, 31))
    first_day = parse_date_only(date)
    people = [a.strip() for a in attendees.split(";") if a.strip()]
    if not people:
        raise ValueError("Provide at least one attendee.")

    start_idx = day_start_hour * 60 // granularity_minutes
    end_idx = day_end_hour * 60 // granularity_minutes
    needed = duration_minutes // granularity_minutes
    if duration_minutes % granularity_minutes != 0:
        needed += 1

    with outlook_session() as ns:
        roster = list(people)
        if include_me:
            try:
                me = ns.CurrentUser.Address
                roster.append(me)
            except Exception:
                pass

        suggestions: list[str] = []
        unknown: set[str] = set()
        resolved_names: list[str] = []
        for d in range(days_to_search):
            if len(suggestions) >= max_suggestions:
                break
            day = first_day + timedelta(days=d)
            combined, day_unknown, names = _free_slots_for_day(
                ns, roster, day, granularity_minutes, start_idx, end_idx
            )
            unknown |= day_unknown
            for n in names:
                if n not in resolved_names:
                    resolved_names.append(n)
            idx = start_idx
            while idx <= end_idx - needed and len(suggestions) < max_suggestions:
                if all(combined[idx : idx + needed]):
                    s = slot_time(day, idx, granularity_minutes)
                    e = slot_time(day, idx + needed, granularity_minutes)
                    suggestions.append(
                        f"  {s.strftime('%a %Y-%m-%d %H:%M')}–{e.strftime('%H:%M')}"
                    )
                    idx += needed
                else:
                    idx += 1

        window = f"{date} for {days_to_search} day(s)"
        header = (
            f"Common {duration_minutes}-minute openings for "
            f"{', '.join(resolved_names)} ({window}, "
            f"{day_start_hour:02d}:00–{day_end_hour:02d}:00 each day):"
        )
        body = (
            "\n".join(suggestions)
            if suggestions
            else "  None found — everyone is busy or free/busy data is unavailable."
        )
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
# Meetings & invitations
# ---------------------------------------------------------------------------


@mcp.tool()
def create_teams_meeting(
    subject: str,
    start: str,
    duration_minutes: int = 30,
    attendees: str = "",
    location: str = "",
    body: str = "",
    timezone_name: str = "",
    save_as_draft: bool = False,
) -> str:
    """Create and send a meeting invite, intended as a Microsoft Teams meeting.

    IMPORTANT — how the Teams link is added: COM cannot inject a Teams join
    link directly (the link comes from the Teams Meeting Add-in / service, not
    the Outlook Object Model). This tool relies on the mailbox/Outlook setting
    "Add online meeting to all meetings" (Outlook > File > Options > Calendar).
    When that setting is ON, meetings created and sent here automatically become
    Teams meetings with a join link. When it is OFF, this sends an ordinary
    (non-online) meeting invite — the return message will remind you to enable
    the setting or add the Teams link manually.

    Args:
        subject: Meeting title.
        start: Start time, "YYYY-MM-DD HH:MM" (wall-clock in timezone_name).
        duration_minutes: Length in minutes.
        attendees: Semicolon-separated attendee addresses (required to invite).
        location: Optional location (ignored for online-only meetings).
        body: Optional agenda/description.
        timezone_name: Time zone for `start` (IANA / Windows ID / abbreviation);
            omit to use Outlook's profile zone.
        save_as_draft: Save to Drafts instead of sending.
    """
    start_dt = parse_when(start)
    with outlook_session() as ns:
        app = ns.Application
        appt = app.CreateItem(ITEM_APPOINTMENT)
        appt.Subject = subject
        tzinfo: dict[str, str] = {}
        if timezone_name:
            tzinfo = apply_timezone(app, appt, start_dt, timezone_name)
        appt.Start = start_dt
        appt.Duration = duration_minutes
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.MeetingStatus = MEETING_MEETING
        resolved: list[dict[str, str]] = []
        if split_recipients(attendees):
            resolved = add_and_resolve_recipients(
                appt, [(attendees, "Required", 1)], allow_unresolved=False
            )

        # Best-effort probe for whether it became an online/Teams meeting.
        is_online = False
        try:
            is_online = bool(appt.IsOnlineMeeting)  # not present on all builds
        except Exception:
            is_online = False

        if save_as_draft:
            appt.Save()
            action = "Meeting draft saved"
        else:
            appt.Send()
            action = "Meeting invite sent"

        note = (
            ""
            if is_online
            else (
                "\nNote: could not confirm a Teams link was attached. If this "
                "meeting isn't a Teams meeting, enable Outlook > File > Options "
                "> Calendar > 'Add online meeting to all meetings', or add the "
                "Teams link manually — COM cannot inject it."
            )
        )
        who = (
            f"\nAttendees:\n{format_recipient_lines(resolved)}" if resolved else ""
        )
        return (
            f"{action}: {subject!r} at {start} ({duration_minutes} min)"
            f"{_tz_note(tzinfo)}{who}{note}"
        )


@mcp.tool()
def create_recurring_meeting(
    subject: str,
    start: str,
    recurrence: str,
    duration_minutes: int = 30,
    interval: int = 1,
    days: str = "",
    count: int = 0,
    until: str = "",
    attendees: str = "",
    location: str = "",
    body: str = "",
    timezone_name: str = "",
    save_as_draft: bool = False,
) -> str:
    """Create and send a recurring meeting invitation.

    Same Teams-link caveat as create_teams_meeting: a Teams join link is only
    added if your mailbox has "Add online meeting to all meetings" enabled — COM
    cannot inject it. Recipients receive the whole series; when they accept, they
    accept the series.

    Setting timezone_name is strongly recommended for recurring series: it pins
    the wall-clock time to a real zone so occurrences stay correct across a
    daylight-saving transition instead of drifting by an hour.

    Args:
        subject: Meeting title.
        start: First occurrence start, "YYYY-MM-DD HH:MM" (in timezone_name).
        recurrence: "daily", "weekly", "monthly", or "yearly".
        duration_minutes: Length of each occurrence.
        interval: Repeat every N units (e.g. interval=2 + weekly = fortnightly).
        days: For weekly recurrence, which weekdays, e.g. "Mon,Wed,Fri".
            Defaults to the weekday of `start` if omitted.
        count: End the series after this many occurrences (0 = use `until` or
            no end date).
        until: End the series on this date, "YYYY-MM-DD" (ignored if count > 0).
        attendees: Semicolon-separated attendee addresses.
        location: Optional location.
        body: Optional agenda/description.
        timezone_name: Time zone for the series (IANA / Windows ID /
            abbreviation). Omit to use Outlook's profile zone (not recommended
            for recurring meetings that cross a DST boundary).
        save_as_draft: Save to Drafts instead of sending.
    """
    start_dt = parse_when(start)
    with outlook_session() as ns:
        app = ns.Application
        appt = app.CreateItem(ITEM_APPOINTMENT)
        appt.Subject = subject
        tzinfo: dict[str, str] = {}
        if timezone_name:
            tzinfo = apply_timezone(app, appt, start_dt, timezone_name)
        appt.Start = start_dt
        appt.Duration = duration_minutes
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.MeetingStatus = MEETING_MEETING
        resolved: list[dict[str, str]] = []
        if split_recipients(attendees):
            resolved = add_and_resolve_recipients(
                appt, [(attendees, "Required", 1)], allow_unresolved=False
            )

        # Configure recurrence last — setting Start after this resets it.
        summary = apply_recurrence(
            appt, recurrence, interval, days, count, until, start_dt, duration_minutes
        )

        if save_as_draft:
            appt.Save()
            action = "Recurring meeting draft saved"
        else:
            appt.Send()
            action = "Recurring meeting invite sent"
        who = (
            f"\nAttendees:\n{format_recipient_lines(resolved)}" if resolved else ""
        )
        return (
            f"{action}: {subject!r} ({summary}), first at {start}"
            f"{_tz_note(tzinfo)}{who}"
        )


def _calendar_conflicts(ns: Any, start: Any, end: Any, ignore_subject: str) -> list[str]:
    """Return subjects of existing calendar events overlapping [start, end)."""
    conflicts: list[str] = []
    try:
        items = ns.GetDefaultFolder(FOLDER_CALENDAR).Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        items = items.Restrict(
            f"[Start] < '{restrict_date(end)}' AND [End] > '{restrict_date(start)}'"
        )
        for appt in items:
            if appt.Subject != ignore_subject:
                conflicts.append(f"{appt.Subject} ({fmt_dt(appt.Start)})")
            if len(conflicts) >= 5:
                break
    except Exception:
        pass
    return conflicts


@mcp.tool()
def list_meeting_invitations(
    count: int = 20, status: str = "pending", show_conflicts: bool = True
) -> str:
    """List meeting invitations in your Inbox, with response state and conflicts.

    Args:
        count: Maximum invitations to return (1-100).
        status: Filter by your response state — "pending", "accepted",
            "tentative", "declined", or "all".
        show_conflicts: Flag other calendar events that overlap each invite.
    """
    count = max(1, min(count, 100))
    status = status.strip().lower()
    valid = {"pending", "accepted", "tentative", "declined", "all"}
    if status not in valid:
        raise ValueError(f"status must be one of: {', '.join(sorted(valid))}.")
    # Map our filter word to the OlResponseStatus of the associated appointment.
    want = {
        "pending": {0, 5},  # None / NotResponded
        "accepted": {3},
        "tentative": {2},
        "declined": {4},
    }
    with outlook_session() as ns:
        items = resolve_folder(ns, "Inbox").Items
        items.Sort("[ReceivedTime]", True)
        results = []
        for item in items:
            if getattr(item, "Class", None) != MEETING_REQUEST_CLASS:
                continue
            when = "(unknown time)"
            location = ""
            resp_code = 0
            conflicts: list[str] = []
            try:
                appt = item.GetAssociatedAppointment(False)
                when = f"{fmt_dt(appt.Start)} -> {fmt_dt(appt.End)}"
                location = appt.Location or ""
                resp_code = getattr(appt, "ResponseStatus", 0)
                if show_conflicts:
                    conflicts = _calendar_conflicts(
                        ns, appt.Start, appt.End, appt.Subject
                    )
            except Exception:
                pass
            if status != "all" and resp_code not in want[status]:
                continue
            loc = f"  |  {location}" if location else ""
            conf = (
                f"\n   CONFLICTS: {'; '.join(conflicts)}" if conflicts else ""
            )
            results.append(
                f"{len(results) + 1}. {item.Subject}\n"
                f"   Organizer: {item.SenderName}\n"
                f"   When: {when}{loc}\n"
                f"   My status: {RESPONSE_STATUS.get(resp_code, '?')}{conf}\n"
                f"   entry_id: {item.EntryID}"
            )
            if len(results) >= count:
                break
        return "\n\n".join(results) if results else f"No {status} meeting invitations."


@mcp.tool()
def respond_to_invitation(
    entry_id: str,
    response: str,
    message: str = "",
    send_response: bool = True,
) -> str:
    """Accept, tentatively accept, or decline a meeting invitation.

    NOTE: "Propose new time" is NOT available through COM — the Outlook Object
    Model has no propose-new-time method. To suggest a different time, decline
    with a message here (or reply by email) and let the organizer reschedule.

    Args:
        entry_id: EntryID of the meeting invitation (from list_meeting_invitations)
            or of an appointment already on your calendar.
        response: One of "accept", "tentative", or "decline".
        message: Optional note included with your response to the organizer.
        send_response: If True, send the response to the organizer; if False,
            respond without notifying them (Outlook's "Do Not Send Response").
    """
    key = response.strip().lower()
    if key not in RESPONSE_MAP:
        raise ValueError('response must be one of: "accept", "tentative", "decline".')
    code = RESPONSE_MAP[key]
    with outlook_session() as ns:
        obj = get_item(ns, entry_id)
        if hasattr(obj, "GetAssociatedAppointment"):
            appt = obj.GetAssociatedAppointment(True)
        else:
            appt = obj  # already an AppointmentItem
        subject = appt.Subject
        resp = appt.Respond(code, True, False)
        verb = {
            RESPONSE_ACCEPTED: "Accepted",
            RESPONSE_TENTATIVE: "Tentatively accepted",
            RESPONSE_DECLINED: "Declined",
        }[code]
        if resp is None:
            return f"{verb} {subject!r} (no response object returned)."
        if send_response:
            if message:
                resp.Body = message + "\n\n" + (resp.Body or "")
            resp.Send()
            return f"{verb} {subject!r} and sent a response to the organizer."
        return f"{verb} {subject!r} without sending a response."


@mcp.tool()
def list_shared_calendar(person: str, days_ahead: int = 7, days_back: int = 0) -> str:
    """Read the actual meetings on another person's calendar (if shared to you).

    Unlike check_availability (which shows only free/busy blocks), this shows
    real meeting subjects and times — but only works if that person has shared
    their calendar with you at Reviewer permission or above. Without permission,
    Outlook denies access and this returns an error; use check_availability for
    plain free/busy instead.

    Args:
        person: Name or email of the calendar owner.
        days_ahead: Days into the future to include.
        days_back: Days into the past to include.
    """
    start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(
        days=days_back
    )
    end = start + timedelta(days=days_back + days_ahead + 1)
    with outlook_session() as ns:
        recip = resolve_recipient(ns, person)
        try:
            folder = ns.GetSharedDefaultFolder(recip, FOLDER_CALENDAR)
        except Exception as exc:
            raise ValueError(
                f"Cannot open {recip.Name}'s calendar — they likely have not "
                f"shared it with you (Reviewer permission needed). {exc}"
            ) from exc
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        items = items.Restrict(
            f"[Start] >= '{restrict_date(start)}' AND [Start] < '{restrict_date(end)}'"
        )
        lines = []
        for appt in items:
            if len(lines) >= 100:
                lines.append("[... truncated at 100 events ...]")
                break
            location = f" @ {appt.Location}" if appt.Location else ""
            lines.append(
                f"- {appt.Subject}{location}\n"
                f"  {fmt_dt(appt.Start)} -> {fmt_dt(appt.End)}"
            )
        header = f"{recip.Name}'s calendar:"
        return f"{header}\n" + ("\n".join(lines) if lines else "  No events in this window.")


# ---------------------------------------------------------------------------
# Calendar event lifecycle (get / search / update / cancel / delete)
# ---------------------------------------------------------------------------


def _require_appointment(ns: Any, entry_id: str) -> Any:
    item = get_item(ns, entry_id)
    if getattr(item, "Class", None) not in (APPOINTMENT_ITEM_CLASS, None):
        raise ValueError(
            f"entry_id {entry_id!r} is not a calendar appointment "
            f"(item class {getattr(item, 'Class', '?')})."
        )
    return item


@mcp.tool()
def get_calendar_event(entry_id: str) -> str:
    """Get full details of a calendar event: organizer, attendees and their
    response status, recurrence, Teams/online join links, importance,
    sensitivity, categories, attachments, and body.

    Args:
        entry_id: EntryID of the appointment (from list_calendar_events etc.).
    """
    with outlook_session() as ns:
        return appointment_details(_require_appointment(ns, entry_id))


@mcp.tool()
def search_calendar_events(
    query: str = "",
    days_ahead: int = 365,
    days_back: int = 0,
    organizer: str = "",
    location: str = "",
    category: str = "",
    teams_only: bool = False,
    recurring_only: bool = False,
    has_attachments: bool = False,
    count: int = 50,
) -> str:
    """Search your calendar with filters over a date window.

    Args:
        query: Text to match in the subject (case-insensitive substring).
        days_ahead: Days into the future to include.
        days_back: Days into the past to include.
        organizer: Only events whose organizer name/address contains this text.
        location: Only events whose location contains this text.
        category: Only events tagged with this category (substring).
        teams_only: Only events that contain an online/Teams join link.
        recurring_only: Only recurring events.
        has_attachments: Only events that have attachments.
        count: Maximum results (1-200).
    """
    count = max(1, min(count, 200))
    start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(
        days=days_back
    )
    end = start + timedelta(days=days_back + days_ahead + 1)
    ql, orgl, locl, catl = (
        query.lower(), organizer.lower(), location.lower(), category.lower()
    )
    with outlook_session() as ns:
        items = ns.GetDefaultFolder(FOLDER_CALENDAR).Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        items = items.Restrict(
            f"[Start] >= '{restrict_date(start)}' AND [Start] < '{restrict_date(end)}'"
        )
        results = []
        for appt in items:
            try:
                if ql and ql not in (appt.Subject or "").lower():
                    continue
                if orgl and orgl not in (getattr(appt, "Organizer", "") or "").lower():
                    continue
                if locl and locl not in (getattr(appt, "Location", "") or "").lower():
                    continue
                if catl and catl not in (getattr(appt, "Categories", "") or "").lower():
                    continue
                if recurring_only and not appt.IsRecurring:
                    continue
                if has_attachments and not appt.Attachments.Count:
                    continue
                if teams_only and not extract_meeting_links(getattr(appt, "Body", "")):
                    continue
            except Exception:
                continue
            loc = f" @ {appt.Location}" if getattr(appt, "Location", "") else ""
            tflag = " [Teams]" if extract_meeting_links(getattr(appt, "Body", "")) else ""
            rflag = " [recurring]" if appt.IsRecurring else ""
            results.append(
                f"{len(results) + 1}. {appt.Subject}{loc}{tflag}{rflag}\n"
                f"   {fmt_dt(appt.Start)} -> {fmt_dt(appt.End)}\n"
                f"   entry_id: {appt.EntryID}"
            )
            if len(results) >= count:
                break
        return "\n\n".join(results) if results else "No matching events found."


@mcp.tool()
def update_calendar_event(
    entry_id: str,
    subject: str | None = None,
    start: str | None = None,
    duration_minutes: int | None = None,
    location: str | None = None,
    body: str | None = None,
    timezone_name: str = "",
    add_attendees: str = "",
    remove_attendees: str = "",
    send_update: bool = True,
) -> str:
    """Update an existing calendar event or meeting.

    For a recurring series, this edits the WHOLE series (use
    update_single_occurrence for one instance). Only the fields you pass are
    changed. If the event has attendees, send_update controls whether an updated
    invitation is sent.

    Args:
        entry_id: EntryID of the appointment.
        subject: New subject (optional).
        start: New start "YYYY-MM-DD HH:MM" (optional).
        duration_minutes: New duration (optional).
        location: New location (optional).
        body: New body/agenda (optional).
        timezone_name: Time zone for `start` (IANA / Windows ID / abbreviation).
        add_attendees: Semicolon-separated addresses to add as required.
        remove_attendees: Semicolon-separated addresses/names to remove.
        send_update: Send an updated invite to attendees (default True).
    """
    with outlook_session() as ns:
        app = ns.Application
        appt = _require_appointment(ns, entry_id)
        changed = []
        if subject is not None:
            appt.Subject = subject
            changed.append("subject")
        tzinfo: dict[str, str] = {}
        if start is not None:
            start_dt = parse_when(start)
            if timezone_name:
                tzinfo = apply_timezone(app, appt, start_dt, timezone_name)
            appt.Start = start_dt
            changed.append("start")
        if duration_minutes is not None:
            appt.Duration = duration_minutes
            changed.append("duration")
        if location is not None:
            appt.Location = location
            changed.append("location")
        if body is not None:
            appt.Body = body
            changed.append("body")
        if remove_attendees.strip():
            targets = {a.strip().lower() for a in split_recipients(remove_attendees)}
            # Iterate backwards; Recipients is 1-based in COM.
            for i in range(appt.Recipients.Count, 0, -1):
                r = appt.Recipients.Item(i)
                if (getattr(r, "Name", "") or "").lower() in targets or (
                    recipient_smtp(r).lower() in targets
                ):
                    appt.Recipients.Remove(i)
                    changed.append(f"removed {r.Name}")
        if add_attendees.strip():
            add_and_resolve_recipients(
                appt, [(add_attendees, "Required", 1)], allow_unresolved=False
            )
            changed.append("added attendees")

        has_attendees = appt.Recipients.Count > 0 and appt.MeetingStatus != MEETING_NONMEETING
        if has_attendees and send_update:
            appt.Send()
            how = "updated invite sent to attendees"
        else:
            appt.Save()
            how = "saved (no invite sent)"
        summary = ", ".join(changed) if changed else "no fields"
        return (
            f"Updated {appt.Subject!r}: {summary}; {how}.{_tz_note(tzinfo)}\n"
            f"entry_id: {appt.EntryID}"
        )


@mcp.tool()
def cancel_calendar_event(entry_id: str, notify: bool = True) -> str:
    """Cancel a meeting you organize, notifying attendees.

    Sets the meeting to Canceled and (if notify) sends a cancellation notice,
    then removes it from your calendar. For a single occurrence of a series, use
    cancel_single_occurrence.

    Args:
        entry_id: EntryID of the meeting you organize.
        notify: Send a cancellation notice to attendees (default True).
    """
    with outlook_session() as ns:
        appt = _require_appointment(ns, entry_id)
        subject = appt.Subject
        try:
            appt.MeetingStatus = MEETING_CANCELED
            if notify and appt.Recipients.Count > 0:
                appt.Send()
        except Exception:
            # Non-meeting appointment: nothing to notify, just delete below.
            pass
        appt.Delete()
        note = " and notified attendees" if notify else ""
        return f"Cancelled {subject!r}{note}."


@mcp.tool()
def delete_calendar_event(entry_id: str) -> str:
    """Delete a calendar event without sending any cancellation notice.

    Use for events you don't organize, or to silently remove one. To notify
    attendees of a meeting you own, use cancel_calendar_event instead.
    """
    with outlook_session() as ns:
        appt = _require_appointment(ns, entry_id)
        subject = appt.Subject
        appt.Delete()
        return f"Deleted calendar event {subject!r} (no notice sent)."


def _occurrence(appt: Any, occurrence_date: str) -> Any:
    """Return the AppointmentItem for one occurrence of a recurring series."""
    if not appt.IsRecurring:
        raise ValueError("This event is not recurring; use the non-occurrence tools.")
    rp = appt.GetRecurrencePattern()
    occ_dt = parse_when(occurrence_date)
    try:
        return rp.GetOccurrence(occ_dt)
    except Exception as exc:
        raise ValueError(
            f"No occurrence found at {occurrence_date!r}. Pass the exact start "
            "date and time of the occurrence (as shown by list_calendar_events)."
        ) from exc


@mcp.tool()
def update_single_occurrence(
    entry_id: str,
    occurrence_date: str,
    subject: str | None = None,
    start: str | None = None,
    duration_minutes: int | None = None,
    location: str | None = None,
    body: str | None = None,
    send_update: bool = True,
) -> str:
    """Edit ONE occurrence of a recurring meeting, leaving the rest unchanged.

    Args:
        entry_id: EntryID of the recurring master (from list_calendar_events).
        occurrence_date: The occurrence's current start, "YYYY-MM-DD HH:MM".
        subject: New subject for just this occurrence (optional).
        start: New start for this occurrence (optional).
        duration_minutes: New duration for this occurrence (optional).
        location: New location for this occurrence (optional).
        body: New body for this occurrence (optional).
        send_update: Send an updated invite to attendees (default True).
    """
    with outlook_session() as ns:
        master = _require_appointment(ns, entry_id)
        occ = _occurrence(master, occurrence_date)
        changed = []
        if subject is not None:
            occ.Subject = subject
            changed.append("subject")
        if start is not None:
            occ.Start = parse_when(start)
            changed.append("start")
        if duration_minutes is not None:
            occ.Duration = duration_minutes
            changed.append("duration")
        if location is not None:
            occ.Location = location
            changed.append("location")
        if body is not None:
            occ.Body = body
            changed.append("body")
        if occ.Recipients.Count > 0 and send_update:
            occ.Send()
            how = "updated invite sent"
        else:
            occ.Save()
            how = "saved"
        summary = ", ".join(changed) if changed else "no fields"
        return f"Updated occurrence on {occurrence_date}: {summary}; {how}."


@mcp.tool()
def cancel_single_occurrence(entry_id: str, occurrence_date: str) -> str:
    """Delete/cancel ONE occurrence of a recurring meeting (an exception).

    Args:
        entry_id: EntryID of the recurring master.
        occurrence_date: The occurrence's start, "YYYY-MM-DD HH:MM".
    """
    with outlook_session() as ns:
        master = _require_appointment(ns, entry_id)
        subject = master.Subject
        occ = _occurrence(master, occurrence_date)
        occ.Delete()
        return f"Cancelled the {occurrence_date} occurrence of {subject!r}."


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


# ---------------------------------------------------------------------------
# Accounts, diagnostics & health
# ---------------------------------------------------------------------------


@mcp.tool()
def list_accounts() -> str:
    """List the Outlook accounts and stores available in this profile.

    Shows each sending account's SMTP address (use it as the `account` argument
    to send_email for send-as) and each mail store/mailbox (use its name as the
    first path segment in folder arguments, e.g. "you@corp.com/Inbox").
    """
    with outlook_session() as ns:
        app = ns.Application
        lines = ["Accounts (for sending):"]
        try:
            accounts = app.Session.Accounts
            for i in range(1, accounts.Count + 1):
                acct = accounts.Item(i)
                smtp = getattr(acct, "SmtpAddress", "") or "(no smtp)"
                lines.append(f"  - {acct.DisplayName} <{smtp}>")
        except Exception as exc:
            lines.append(f"  (could not enumerate accounts: {exc})")
        lines.append("\nStores (mailboxes / PSTs):")
        try:
            for store_folder in ns.Folders:
                lines.append(f"  - {store_folder.Name}")
        except Exception as exc:
            lines.append(f"  (could not enumerate stores: {exc})")
        return "\n".join(lines)


@mcp.tool()
def outlook_health() -> str:
    """Report Outlook connection status and environment for diagnostics.

    Returns the Outlook version, current user, default store, connection mode,
    and account count — useful for confirming the server can reach a live,
    classic Outlook instance before running other tools.
    """
    if pythoncom is None:
        return (
            "UNAVAILABLE: pywin32 is not installed. This server requires Windows "
            "with classic Outlook desktop."
        )
    with outlook_session() as ns:
        app = ns.Application
        lines = ["Outlook health: OK"]
        for label, getter in (
            ("Version", lambda: app.Version),
            ("Current user", lambda: ns.CurrentUser.Name),
            ("Default store", lambda: ns.DefaultStore.DisplayName),
            (
                "Exchange mode",
                lambda: {
                    100: "No Exchange",
                    200: "Offline",
                    300: "Online",
                    400: "Disconnected",
                    500: "Connected headers only",
                }.get(ns.ExchangeConnectionMode, str(ns.ExchangeConnectionMode)),
            ),
            ("Accounts", lambda: app.Session.Accounts.Count),
        ):
            try:
                lines.append(f"  {label}: {getter()}")
            except Exception as exc:
                lines.append(f"  {label}: (unavailable: {exc})")
        return "\n".join(lines)


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
