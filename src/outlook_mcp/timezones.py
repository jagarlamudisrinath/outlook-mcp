"""Timezone parsing and conversion helpers (pure stdlib, no COM).

Kept separate from server.py so the DST-sensitive logic can be unit-tested on
any platform — the Outlook COM layer cannot run off Windows, but this can.

Outlook's COM API identifies time zones by **Windows** IDs (e.g.
"India Standard Time"), while DST-correct arithmetic in Python needs **IANA**
names (e.g. "Asia/Kolkata"). This module bridges the two and produces
timezone-aware datetimes so callers can report both local and UTC times.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

DATETIME_FORMATS = ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d")

# Curated Windows time-zone ID -> IANA name map covering common business zones.
# Not exhaustive; unknown-but-valid IANA names still work (Windows ID stays None
# and the COM layer falls back to the profile zone while UTC is still correct).
WINDOWS_TO_IANA: dict[str, str] = {
    "UTC": "UTC",
    "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik",
    "W. Europe Standard Time": "Europe/Berlin",
    "Central Europe Standard Time": "Europe/Budapest",
    "Central European Standard Time": "Europe/Warsaw",
    "Romance Standard Time": "Europe/Paris",
    "E. Europe Standard Time": "Europe/Bucharest",
    "India Standard Time": "Asia/Kolkata",
    "Sri Lanka Standard Time": "Asia/Colombo",
    "Pakistan Standard Time": "Asia/Karachi",
    "Bangladesh Standard Time": "Asia/Dhaka",
    "Nepal Standard Time": "Asia/Kathmandu",
    "China Standard Time": "Asia/Shanghai",
    "Tokyo Standard Time": "Asia/Tokyo",
    "Korea Standard Time": "Asia/Seoul",
    "Singapore Standard Time": "Asia/Singapore",
    "SE Asia Standard Time": "Asia/Bangkok",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "W. Australia Standard Time": "Australia/Perth",
    "New Zealand Standard Time": "Pacific/Auckland",
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver",
    "US Mountain Standard Time": "America/Phoenix",
    "Pacific Standard Time": "America/Los_Angeles",
    "Alaskan Standard Time": "America/Anchorage",
    "Hawaiian Standard Time": "Pacific/Honolulu",
    "Central Standard Time (Mexico)": "America/Mexico_City",
    "Canada Central Standard Time": "America/Regina",
    "SA Pacific Standard Time": "America/Bogota",
    "E. South America Standard Time": "America/Sao_Paulo",
    "Arabian Standard Time": "Asia/Dubai",
    "Arab Standard Time": "Asia/Riyadh",
    "Israel Standard Time": "Asia/Jerusalem",
    "Turkey Standard Time": "Europe/Istanbul",
    "Russian Standard Time": "Europe/Moscow",
    "South Africa Standard Time": "Africa/Johannesburg",
    "E. Africa Standard Time": "Africa/Nairobi",
    "W. Central Africa Standard Time": "Africa/Lagos",
    "Egypt Standard Time": "Africa/Cairo",
}

IANA_TO_WINDOWS: dict[str, str] = {iana: win for win, iana in WINDOWS_TO_IANA.items()}

# Common abbreviations users type. Ambiguous ones resolve to the most common
# business meaning (documented in the tool help). IST -> India, not Israel/Irish.
ALIASES: dict[str, str] = {
    "UTC": "UTC",
    "GMT": "Europe/London",
    "BST": "Europe/London",
    "IST": "Asia/Kolkata",
    "PST": "America/Los_Angeles",
    "PDT": "America/Los_Angeles",
    "PT": "America/Los_Angeles",
    "MST": "America/Denver",
    "MT": "America/Denver",
    "CST": "America/Chicago",
    "CT": "America/Chicago",
    "EST": "America/New_York",
    "EDT": "America/New_York",
    "ET": "America/New_York",
    "CET": "Europe/Berlin",
    "CEST": "Europe/Berlin",
}


@dataclass(frozen=True)
class ResolvedTimeZone:
    """A time zone resolved to both naming schemes plus a live tzinfo."""

    iana: str
    windows: str | None
    tzinfo: ZoneInfo

    @property
    def label(self) -> str:
        return self.iana if self.windows is None else f"{self.iana} ({self.windows})"


def parse_naive(value: str) -> datetime:
    """Parse a date/time string into a naive datetime (no timezone attached)."""
    value = value.strip()
    for fmt in DATETIME_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise ValueError(
        f"Could not parse datetime {value!r}. Use 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD'."
    )


def resolve_timezone(tz: str) -> ResolvedTimeZone:
    """Resolve a Windows ID, IANA name, or common abbreviation to a time zone.

    Raises ValueError if the zone cannot be identified.
    """
    tz = tz.strip()
    if not tz:
        raise ValueError("Empty timezone.")

    # 1) Exact Windows ID.
    if tz in WINDOWS_TO_IANA:
        iana = WINDOWS_TO_IANA[tz]
        return ResolvedTimeZone(iana, tz, ZoneInfo(iana))

    # 2) Abbreviation.
    upper = tz.upper()
    if upper in ALIASES:
        iana = ALIASES[upper]
        return ResolvedTimeZone(iana, IANA_TO_WINDOWS.get(iana), ZoneInfo(iana))

    # 3) IANA name (validated by ZoneInfo).
    try:
        info = ZoneInfo(tz)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        raise ValueError(
            f"Unrecognized timezone {tz!r}. Use an IANA name like "
            "'Asia/Kolkata', a Windows ID like 'India Standard Time', or a "
            "common abbreviation like 'IST'."
        ) from exc
    return ResolvedTimeZone(tz, IANA_TO_WINDOWS.get(tz), info)


def localize(naive: datetime, rtz: ResolvedTimeZone) -> datetime:
    """Attach a resolved time zone to a naive wall-clock datetime."""
    if naive.tzinfo is not None:
        raise ValueError("localize() expects a naive datetime.")
    return naive.replace(tzinfo=rtz.tzinfo)


def to_utc(aware: datetime) -> datetime:
    """Convert a timezone-aware datetime to UTC."""
    if aware.tzinfo is None:
        raise ValueError("to_utc() expects a timezone-aware datetime.")
    return aware.astimezone(timezone.utc)


def fmt(dt: datetime) -> str:
    """Format a datetime as 'YYYY-MM-DD HH:MM' (with offset if aware)."""
    base = dt.strftime("%Y-%m-%d %H:%M")
    if dt.tzinfo is not None:
        return f"{base} {dt.strftime('%z')}"
    return base
