"""Tests for the timezone module — DST-correct, runnable on any platform."""

from __future__ import annotations

from datetime import datetime

import pytest

from outlook_mcp import timezones as tz


def test_parse_naive_formats():
    assert tz.parse_naive("2026-07-16 14:30") == datetime(2026, 7, 16, 14, 30)
    assert tz.parse_naive("2026-07-16T14:30") == datetime(2026, 7, 16, 14, 30)
    assert tz.parse_naive("2026-07-16") == datetime(2026, 7, 16, 0, 0)
    assert tz.parse_naive("  2026-07-16 09:05  ") == datetime(2026, 7, 16, 9, 5)


def test_parse_naive_rejects_garbage():
    with pytest.raises(ValueError):
        tz.parse_naive("not a date")


def test_resolve_windows_id():
    r = tz.resolve_timezone("India Standard Time")
    assert r.iana == "Asia/Kolkata"
    assert r.windows == "India Standard Time"


def test_resolve_iana_name():
    r = tz.resolve_timezone("America/New_York")
    assert r.iana == "America/New_York"
    assert r.windows == "Eastern Standard Time"


def test_resolve_abbreviation_ist_is_india():
    r = tz.resolve_timezone("IST")
    assert r.iana == "Asia/Kolkata"
    assert r.windows == "India Standard Time"


def test_resolve_unknown_iana_has_no_windows_id():
    # A valid IANA zone we don't map keeps windows=None (COM falls back).
    r = tz.resolve_timezone("America/Argentina/Ushuaia")
    assert r.iana == "America/Argentina/Ushuaia"
    assert r.windows is None


def test_resolve_rejects_nonsense():
    with pytest.raises(ValueError):
        tz.resolve_timezone("Middle/Earth")
    with pytest.raises(ValueError):
        tz.resolve_timezone("")


def test_utc_conversion_standard_time():
    r = tz.resolve_timezone("America/Los_Angeles")
    aware = tz.localize(tz.parse_naive("2026-01-15 12:00"), r)
    assert tz.to_utc(aware).strftime("%Y-%m-%d %H:%M") == "2026-01-15 20:00"  # PST -8


def test_utc_conversion_daylight_time():
    r = tz.resolve_timezone("America/Los_Angeles")
    aware = tz.localize(tz.parse_naive("2026-07-15 12:00"), r)
    assert tz.to_utc(aware).strftime("%Y-%m-%d %H:%M") == "2026-07-15 19:00"  # PDT -7


def test_dst_boundary_same_walltime_different_utc():
    """The reviewer's bug: identical wall-clock time yields different UTC across
    a DST transition. A recurring 12:00 meeting must NOT drift in UTC terms
    naively — this documents why an explicit zone is required."""
    r = tz.resolve_timezone("America/New_York")
    before = tz.to_utc(tz.localize(tz.parse_naive("2026-03-07 12:00"), r))  # EST
    after = tz.to_utc(tz.localize(tz.parse_naive("2026-03-09 12:00"), r))  # EDT
    assert before.strftime("%H:%M") == "17:00"
    assert after.strftime("%H:%M") == "16:00"


def test_no_dst_zone_stable():
    r = tz.resolve_timezone("Asia/Kolkata")
    jan = tz.to_utc(tz.localize(tz.parse_naive("2026-01-15 12:00"), r))
    jul = tz.to_utc(tz.localize(tz.parse_naive("2026-07-15 12:00"), r))
    assert jan.strftime("%H:%M") == jul.strftime("%H:%M") == "06:30"


def test_localize_rejects_aware():
    r = tz.resolve_timezone("UTC")
    aware = tz.localize(tz.parse_naive("2026-07-16 10:00"), r)
    with pytest.raises(ValueError):
        tz.localize(aware, r)


def test_to_utc_rejects_naive():
    with pytest.raises(ValueError):
        tz.to_utc(tz.parse_naive("2026-07-16 10:00"))


def test_fmt_includes_offset_for_aware():
    r = tz.resolve_timezone("Asia/Kolkata")
    aware = tz.localize(tz.parse_naive("2026-07-16 10:00"), r)
    assert tz.fmt(aware) == "2026-07-16 10:00 +0530"
    assert tz.fmt(tz.parse_naive("2026-07-16 10:00")) == "2026-07-16 10:00"
