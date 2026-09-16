from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.event_mapper import to_event

SAO_PAULO = ZoneInfo("America/Sao_Paulo")


def graph_time(wall_clock: str, zone: str = "America/Sao_Paulo") -> dict[str, str]:
    return {"dateTime": wall_clock, "timeZone": zone}


def graph_event(**overrides: Any) -> dict[str, Any]:
    event: dict[str, Any] = {
        "id": "AAMkAGI2",
        "subject": "Planning",
        "start": graph_time("2026-09-15T09:00:00.0000000"),
        "end": graph_time("2026-09-15T10:00:00.0000000"),
        "isAllDay": False,
        "location": {"displayName": "Room 1"},
        "organizer": {"emailAddress": {"name": "Ana Lima", "address": "ana@example.com"}},
    }
    return event | overrides


def test_maps_every_field_of_the_entity() -> None:
    event = to_event(graph_event())

    assert event.id == "AAMkAGI2"
    assert event.subject == "Planning"
    assert event.start == datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
    assert event.end == datetime(2026, 9, 15, 10, tzinfo=SAO_PAULO)
    assert event.is_all_day is False
    assert event.location == "Room 1"
    assert event.organizer.address == "ana@example.com"
    assert event.organizer.display_name == "Ana Lima"


def test_attaches_the_zone_graph_names_to_each_time() -> None:
    event = to_event(graph_event(start=graph_time("2026-09-15T09:00:00", "Europe/Lisbon")))

    assert event.start.tzinfo == ZoneInfo("Europe/Lisbon")
    assert event.end.tzinfo == SAO_PAULO


def test_reads_the_seven_fractional_digits_graph_emits() -> None:
    event = to_event(graph_event(start=graph_time("2026-09-15T09:00:00.1234567")))

    assert event.start.microsecond == 123456


def test_maps_an_all_day_event_as_midnight_to_midnight() -> None:
    event = to_event(
        graph_event(
            start=graph_time("2026-09-15T00:00:00.0000000"),
            end=graph_time("2026-09-16T00:00:00.0000000"),
            isAllDay=True,
        )
    )

    assert event.is_all_day is True
    assert event.start == datetime(2026, 9, 15, tzinfo=SAO_PAULO)
    assert event.end == datetime(2026, 9, 16, tzinfo=SAO_PAULO)


def test_keeps_a_time_that_already_carries_an_offset_as_the_same_instant() -> None:
    event = to_event(graph_event(start=graph_time("2026-09-15T12:00:00Z")))

    assert event.start == datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)


def test_tolerates_a_missing_location_and_organizer() -> None:
    payload = graph_event()
    del payload["location"]
    del payload["organizer"]

    event = to_event(payload)

    assert event.location == ""
    assert event.organizer.address == ""


def test_maps_an_absent_subject_to_empty() -> None:
    assert to_event(graph_event(subject=None)).subject == ""


def test_raises_when_the_zone_is_not_an_iana_name() -> None:
    with pytest.raises(GraphResponseError, match="timeZone"):
        to_event(graph_event(start=graph_time("2026-09-15T09:00:00", "Pacific Standard Time")))


def test_raises_when_a_time_is_not_iso_8601() -> None:
    with pytest.raises(GraphResponseError, match="ISO 8601"):
        to_event(graph_event(end=graph_time("tomorrow-ish")))


@pytest.mark.parametrize("field", ["id", "start", "end", "isAllDay"])
def test_raises_when_a_required_field_is_missing(field: str) -> None:
    payload = graph_event()
    del payload[field]

    with pytest.raises(GraphResponseError, match=field):
        to_event(payload)


def test_maps_a_missing_body_to_empty() -> None:
    """A listing never selects body, so a listed event carries none; this must not raise."""
    assert to_event(graph_event()).body == ""


def test_reads_a_plain_text_body() -> None:
    event = to_event(graph_event(body={"contentType": "text", "content": "Bring the deck."}))

    assert event.body == "Bring the deck."


def test_raises_when_the_body_is_html_instead_of_text() -> None:
    with pytest.raises(GraphResponseError, match="html"):
        to_event(graph_event(body={"contentType": "html", "content": "<p>Bring it</p>"}))
