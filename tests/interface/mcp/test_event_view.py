import json
from datetime import datetime
from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.interface.mcp.event_view import EventView

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
AN_EVENT = Event(
    id="AAMkAGI2",
    subject="Planning",
    start=datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO),
    end=datetime(2026, 9, 15, 10, tzinfo=SAO_PAULO),
    is_all_day=False,
    location="Room 1",
    organizer=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
)


def test_projects_every_field_of_the_entity() -> None:
    view = EventView.from_event(AN_EVENT)

    assert view.id == "AAMkAGI2"
    assert view.subject == "Planning"
    assert view.start == datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
    assert view.end == datetime(2026, 9, 15, 10, tzinfo=SAO_PAULO)
    assert view.is_all_day is False
    assert view.location == "Room 1"
    assert view.organizer_address == "ana@example.com"
    assert view.organizer_name == "Ana Lima"


def test_serializes_times_as_iso_8601_with_their_offset() -> None:
    payload = json.loads(EventView.from_event(AN_EVENT).model_dump_json())

    assert payload["start"] == "2026-09-15T09:00:00-03:00"
    assert payload["end"] == "2026-09-15T10:00:00-03:00"
