from datetime import UTC, datetime, timedelta
from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.infrastructure.graph.event_payload import to_event_payload

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
MIDNIGHT = datetime(2026, 9, 15, tzinfo=SAO_PAULO)


def test_sends_wall_clock_times_with_the_resolved_zone_name() -> None:
    payload = to_event_payload(
        NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1)), SAO_PAULO
    )

    assert payload["subject"] == "Planning"
    assert payload["start"] == {"dateTime": "2026-09-15T09:00:00", "timeZone": "America/Sao_Paulo"}
    assert payload["end"] == {"dateTime": "2026-09-15T10:00:00", "timeZone": "America/Sao_Paulo"}
    assert payload["isAllDay"] is False


def test_converts_a_timed_event_given_in_another_zone() -> None:
    payload = to_event_payload(
        NewEvent(
            subject="Call",
            start=datetime(2026, 9, 15, 12, tzinfo=UTC),
            end=datetime(2026, 9, 15, 13, tzinfo=UTC),
        ),
        SAO_PAULO,
    )

    assert payload["start"]["dateTime"] == "2026-09-15T09:00:00"
    assert payload["start"]["timeZone"] == "America/Sao_Paulo"


def test_keeps_the_date_of_an_all_day_event_given_in_another_zone() -> None:
    payload = to_event_payload(
        NewEvent(
            subject="Offsite",
            start=datetime(2026, 9, 15, tzinfo=UTC),
            end=datetime(2026, 9, 16, tzinfo=UTC),
            is_all_day=True,
        ),
        SAO_PAULO,
    )

    assert payload["isAllDay"] is True
    assert payload["start"] == {"dateTime": "2026-09-15T00:00:00", "timeZone": "America/Sao_Paulo"}
    assert payload["end"] == {"dateTime": "2026-09-16T00:00:00", "timeZone": "America/Sao_Paulo"}


def test_omits_the_location_when_there_is_none() -> None:
    payload = to_event_payload(
        NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1)), SAO_PAULO
    )

    assert "location" not in payload


def test_sends_the_location_as_a_display_name() -> None:
    payload = to_event_payload(
        NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1), location="Room 1"),
        SAO_PAULO,
    )

    assert payload["location"] == {"displayName": "Room 1"}


def test_sends_attendees_as_required_with_a_name_only_when_known() -> None:
    payload = to_event_payload(
        NewEvent(
            subject="Planning",
            start=NINE,
            end=NINE + timedelta(hours=1),
            attendees=(
                EmailAddress(address="ana@example.com", display_name="Ana Lima"),
                EmailAddress(address="bo@example.com"),
            ),
        ),
        SAO_PAULO,
    )

    assert payload["attendees"] == [
        {"emailAddress": {"address": "ana@example.com", "name": "Ana Lima"}, "type": "required"},
        {"emailAddress": {"address": "bo@example.com"}, "type": "required"},
    ]
