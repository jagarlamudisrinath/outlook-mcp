"""Tests for server helper logic using mocked COM objects.

These never touch Outlook; they exercise the pure logic (recipient resolution,
recurrence configuration, parsing, interval math) that would otherwise only be
reachable on Windows.
"""

from __future__ import annotations

from datetime import datetime
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from outlook_mcp import server as s


# --- simple parsers -------------------------------------------------------


def test_parse_day_mask_bits():
    assert s.parse_day_mask("Mon,Wed,Fri") == 2 + 8 + 32
    assert s.parse_day_mask("sunday") == 1
    assert s.parse_day_mask("Sat;Sun") == 64 + 1
    assert s.parse_day_mask("") == 0


def test_parse_day_mask_rejects_unknown():
    with pytest.raises(ValueError):
        s.parse_day_mask("Funday")


def test_split_recipients():
    assert s.split_recipients("a@x.com; b@y.com") == ["a@x.com", "b@y.com"]
    assert s.split_recipients("a@x.com;;\nb@y.com ; ") == ["a@x.com", "b@y.com"]
    assert s.split_recipients("   ") == []


def test_dasl_escape():
    assert s.dasl_escape("O'Brien") == "O''Brien"
    assert s.dasl_escape("50%") == "50[%]"


def test_merge_intervals():
    d = datetime
    intervals = [
        (d(2026, 7, 16, 9), d(2026, 7, 16, 10)),
        (d(2026, 7, 16, 9, 30), d(2026, 7, 16, 11)),  # overlaps previous
        (d(2026, 7, 16, 13), d(2026, 7, 16, 14)),  # separate
    ]
    merged = s.merge_intervals(intervals)
    assert merged == [
        (d(2026, 7, 16, 9), d(2026, 7, 16, 11)),
        (d(2026, 7, 16, 13), d(2026, 7, 16, 14)),
    ]
    assert s.merge_intervals([]) == []


# --- recurrence configuration --------------------------------------------


def _mock_appt_with_pattern():
    appt = MagicMock()
    pattern = MagicMock()
    appt.GetRecurrencePattern.return_value = pattern
    return appt, pattern


def test_apply_recurrence_weekly_with_days():
    appt, rp = _mock_appt_with_pattern()
    summary = s.apply_recurrence(
        appt, "weekly", 1, "Mon,Wed", count=6, until="",
        start_dt=datetime(2026, 7, 20, 9, 0), duration_minutes=30,
    )
    assert rp.RecurrenceType == s.RECUR_WEEKLY
    assert rp.Interval == 1
    assert rp.DayOfWeekMask == 2 + 8  # Mon+Wed
    assert rp.Occurrences == 6
    assert "weekly" in summary and "Mon,Wed" in summary and "6 occurrences" in summary


def test_apply_recurrence_monthly_uses_start_day():
    appt, rp = _mock_appt_with_pattern()
    s.apply_recurrence(
        appt, "monthly", 2, "", count=0, until="2026-12-31",
        start_dt=datetime(2026, 7, 15, 9, 0), duration_minutes=60,
    )
    assert rp.RecurrenceType == s.RECUR_MONTHLY
    assert rp.Interval == 2
    assert rp.DayOfMonth == 15
    assert rp.PatternEndDate == datetime(2026, 12, 31, 0, 0)


def test_apply_recurrence_rejects_bad_type():
    appt, _ = _mock_appt_with_pattern()
    with pytest.raises(ValueError):
        s.apply_recurrence(
            appt, "hourly", 1, "", 0, "", datetime(2026, 7, 20, 9), 30
        )


# --- recipient resolution -------------------------------------------------


class FakeRecipient:
    def __init__(self, name, resolved, smtp="", addr_type="SMTP"):
        self.Name = name
        self.Resolved = resolved
        self.Type = None  # set by code under test
        self._smtp = smtp
        exch = SimpleNamespace(PrimarySmtpAddress=smtp) if addr_type == "EX" else None
        self.AddressEntry = SimpleNamespace(
            Type=addr_type,
            Address=smtp,
            GetExchangeUser=lambda: exch,
        )
        self.Address = smtp


class FakeRecipients:
    def __init__(self, table):
        # table maps input address -> FakeRecipient
        self._table = table
        self.added = []

    def Add(self, addr):
        recip = self._table[addr]
        self.added.append(recip)
        return recip

    def ResolveAll(self):
        pass


class FakeItem:
    def __init__(self, table):
        self.Recipients = FakeRecipients(table)


def test_add_and_resolve_all_good():
    table = {
        "alice@corp.com": FakeRecipient("Alice", True, "alice@corp.com"),
        "Bob Smith": FakeRecipient("Bob Smith", True, "bob@corp.com", "EX"),
    }
    item = FakeItem(table)
    resolved = s.add_and_resolve_recipients(
        item, [("alice@corp.com", "To", 1), ("Bob Smith", "CC", 2)],
        allow_unresolved=False,
    )
    assert resolved[0]["smtp"] == "alice@corp.com"
    assert resolved[1]["smtp"] == "bob@corp.com"  # canonical SMTP from Exchange
    assert resolved[1]["type"] == "CC"
    assert table["alice@corp.com"].Type == 1
    assert table["Bob Smith"].Type == 2


def test_add_and_resolve_rejects_unresolved():
    table = {"ghost@nowhere": FakeRecipient("ghost", False)}
    item = FakeItem(table)
    with pytest.raises(ValueError, match="Could not resolve"):
        s.add_and_resolve_recipients(
            item, [("ghost@nowhere", "To", 1)], allow_unresolved=False
        )


def test_add_and_resolve_allow_unresolved():
    table = {"ghost@nowhere": FakeRecipient("ghost", False)}
    item = FakeItem(table)
    resolved = s.add_and_resolve_recipients(
        item, [("ghost@nowhere", "To", 1)], allow_unresolved=True
    )
    assert resolved[0]["resolved"] == "NO"


# --- appointment rendering ------------------------------------------------


def test_appointment_details_renders_key_fields():
    appt = MagicMock()
    appt.Subject = "Sprint Review"
    appt.Organizer = "Alice"
    appt.Start = datetime(2026, 7, 20, 10, 0)
    appt.End = datetime(2026, 7, 20, 11, 0)
    appt.AllDayEvent = False
    appt.Location = "Room 1"
    appt.Duration = 60
    appt.IsRecurring = False
    appt.ResponseStatus = 3  # Accepted
    appt.Importance = 2
    appt.Sensitivity = 0
    appt.Categories = "Work"
    appt.Body = "Agenda: https://teams.microsoft.com/l/meetup-join/xyz please join"
    appt.EntryID = "ENTRY123"
    appt.Attachments.Count = 0
    # One required attendee who accepted.
    recip = SimpleNamespace(
        Type=1,
        MeetingResponseStatus=3,
        Name="Bob",
        AddressEntry=SimpleNamespace(
            Type="SMTP", Address="bob@corp.com", GetExchangeUser=lambda: None
        ),
        Address="bob@corp.com",
    )
    appt.Recipients = [recip]

    out = s.appointment_details(appt)
    assert "Subject: Sprint Review" in out
    assert "Organizer: Alice" in out
    assert "My response: Accepted" in out
    assert "Importance: High" in out
    assert "https://teams.microsoft.com/l/meetup-join/xyz" in out
    assert "[Required] Bob <bob@corp.com> — Accepted" in out
    assert "entry_id: ENTRY123" in out


def test_meeting_attendees_maps_roles_and_status():
    appt = MagicMock()
    appt.Recipients = [
        SimpleNamespace(
            Type=2, MeetingResponseStatus=2, Name="Opt",
            AddressEntry=SimpleNamespace(
                Type="SMTP", Address="o@x.com", GetExchangeUser=lambda: None
            ),
            Address="o@x.com",
        )
    ]
    people = s.meeting_attendees(appt)
    assert people == [
        {"role": "Optional", "name": "Opt", "smtp": "o@x.com", "response": "Tentative"}
    ]
