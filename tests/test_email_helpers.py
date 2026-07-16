"""Tests for email/attachment/search helper logic (no COM)."""

from __future__ import annotations

import pytest

from outlook_mcp import server as s


# --- meeting link extraction ---------------------------------------------


def test_extract_teams_link():
    body = "Join here https://teams.microsoft.com/l/meetup-join/19%3aabc/0?context=x now"
    assert s.extract_meeting_links(body) == [
        "https://teams.microsoft.com/l/meetup-join/19%3aabc/0?context=x"
    ]


def test_extract_dedupes_and_handles_zoom():
    body = (
        "https://acme.zoom.us/j/123456789 and again "
        "https://acme.zoom.us/j/123456789"
    )
    assert s.extract_meeting_links(body) == ["https://acme.zoom.us/j/123456789"]


def test_extract_none():
    assert s.extract_meeting_links("no links here") == []
    assert s.extract_meeting_links("") == []


def test_extract_strips_trailing_punctuation():
    body = "Link: https://teams.microsoft.com/l/meetup-join/abc)."
    assert s.extract_meeting_links(body) == [
        "https://teams.microsoft.com/l/meetup-join/abc"
    ]


# --- filename sanitization -----------------------------------------------


def test_sanitize_strips_path_and_illegal_chars():
    assert s.sanitize_filename("../../etc/pa:ss<w>d.txt") == "pa_ss_w_d.txt"
    assert s.sanitize_filename("C:\\temp\\report.pdf") == "report.pdf"


def test_sanitize_empty_becomes_unnamed():
    assert s.sanitize_filename("") == "unnamed"
    assert s.sanitize_filename("...") == "unnamed"


def test_unique_path_avoids_collision(tmp_path):
    first = s.unique_path(tmp_path, "doc.pdf")
    assert first.name == "doc.pdf"
    first.write_text("x")
    second = s.unique_path(tmp_path, "doc.pdf")
    assert second.name == "doc (1).pdf"
    second.write_text("y")
    third = s.unique_path(tmp_path, "doc.pdf")
    assert third.name == "doc (2).pdf"


# --- DASL builder ---------------------------------------------------------


def test_dasl_requires_a_filter():
    with pytest.raises(ValueError):
        s.build_email_dasl("", "all", "", "", "", "", False, False, False, "")


def test_dasl_query_and_sender():
    dasl = s.build_email_dasl(
        "budget", "subject", "alice", "", "", "", False, False, False, ""
    )
    assert dasl.startswith("@SQL=(")
    assert "urn:schemas:httpmail:subject" in dasl
    assert "%alice%" in dasl
    assert " AND " in dasl


def test_dasl_boolean_and_importance_flags():
    dasl = s.build_email_dasl(
        "", "all", "", "", "", "", True, True, True, "high"
    )
    assert '"urn:schemas:httpmail:read" = 0' in dasl
    assert '"urn:schemas:httpmail:hasattachment" = 1' in dasl
    assert "0x10900003" in dasl  # flag status
    assert '"urn:schemas:httpmail:importance" = 2' in dasl


def test_dasl_date_format():
    dasl = s.build_email_dasl(
        "", "all", "", "", "2026-07-01 09:30", "", False, False, False, ""
    )
    assert "2026/07/01 09:30" in dasl


def test_dasl_rejects_bad_importance():
    with pytest.raises(ValueError):
        s.build_email_dasl("x", "all", "", "", "", "", False, False, False, "urgent")


def test_dasl_rejects_bad_search_in():
    with pytest.raises(ValueError):
        s.build_email_dasl("x", "everywhere", "", "", "", "", False, False, False, "")
