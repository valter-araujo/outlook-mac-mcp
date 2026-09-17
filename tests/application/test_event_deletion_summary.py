from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from outlook_mac_mcp.application.event_deletion_summary import describe_event_for_deletion
from outlook_mac_mcp.application.event_summary import MAX_BODY_IN_SUMMARY
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
ORGANIZER = EmailAddress(address="me@example.com", display_name="Me")


def test_shows_every_field_even_the_empty_ones() -> None:
    event = Event(
        id="AAMkEXISTING",
        subject="Planning",
        start=NINE,
        end=NINE + timedelta(hours=1),
        is_all_day=False,
        location="",
        organizer=ORGANIZER,
    )

    summary = describe_event_for_deletion(event)

    assert 'subject: "Planning"' in summary
    assert "start:" in summary
    assert "end:" in summary
    assert "location: (none)" in summary
    assert "body: (none)" in summary
    assert "attendees: (none)" in summary


def test_shows_every_field_populated() -> None:
    event = Event(
        id="AAMkEXISTING",
        subject="Planning",
        start=NINE,
        end=NINE + timedelta(hours=1),
        is_all_day=False,
        location="Room 1",
        organizer=ORGANIZER,
        body="Bring the deck.",
        attendees=(
            EmailAddress(address="ana@example.com"),
            EmailAddress(address="bo@example.com"),
        ),
    )

    summary = describe_event_for_deletion(event)

    assert 'subject: "Planning"' in summary
    assert 'location: "Room 1"' in summary
    assert 'body: "Bring the deck."' in summary
    assert "attendees: ana@example.com, bo@example.com" in summary


def test_truncates_a_long_body_with_an_explicit_character_count() -> None:
    body = "a" * (MAX_BODY_IN_SUMMARY + 1)
    event = Event(
        id="AAMkEXISTING",
        subject="Planning",
        start=NINE,
        end=NINE + timedelta(hours=1),
        is_all_day=False,
        location="",
        organizer=ORGANIZER,
        body=body,
    )

    summary = describe_event_for_deletion(event)

    assert f'body: "{"a" * MAX_BODY_IN_SUMMARY}…"' in summary
    assert f"{len(body)} characters total" in summary
    assert "truncated" in summary
    # A deletion preview must show every field, but "unabbreviated" governs which
    # fields appear, not whether a body past the display length is still truncated.
    assert body not in summary
