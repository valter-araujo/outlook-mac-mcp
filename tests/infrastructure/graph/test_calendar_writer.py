import json
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import EventNotFoundError
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.domain.sensitivity import Sensitivity
from outlook_mac_mcp.domain.show_as import ShowAs
from outlook_mac_mcp.infrastructure.graph.calendar_writer import GraphCalendarWriter
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError

EVENTS_URL = f"{GRAPH_BASE_URL}/me/events"
AN_EVENT_URL = f"{EVENTS_URL}/AAMkEXISTING"
SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
A_NEW_EVENT = NewEvent(
    subject="Planning", start=NINE, end=NINE + timedelta(hours=1), location="Room 1"
)
A_NEW_EVENT_WITH_BODY = NewEvent(
    subject="Planning", start=NINE, end=NINE + timedelta(hours=1), body="Bring the deck."
)
CREATED: dict[str, Any] = {
    "id": "AAMkNEW",
    "subject": "Planning",
    "start": {"dateTime": "2026-09-15T09:00:00.0000000", "timeZone": "America/Sao_Paulo"},
    "end": {"dateTime": "2026-09-15T10:00:00.0000000", "timeZone": "America/Sao_Paulo"},
    "isAllDay": False,
    "location": {"displayName": "Room 1"},
    "organizer": {"emailAddress": {"name": "Me", "address": "me@example.com"}},
}
CREATED_WITH_BODY: dict[str, Any] = {
    **{key: value for key, value in CREATED.items() if key != "location"},
    "body": {"contentType": "text", "content": "Bring the deck."},
}
EXISTING: dict[str, Any] = {
    "id": "AAMkEXISTING",
    "subject": "Planning",
    "start": {"dateTime": "2026-09-15T09:00:00.0000000", "timeZone": "America/Sao_Paulo"},
    "end": {"dateTime": "2026-09-15T10:00:00.0000000", "timeZone": "America/Sao_Paulo"},
    "isAllDay": False,
    "location": {"displayName": "Room 1"},
    "organizer": {"emailAddress": {"name": "Me", "address": "me@example.com"}},
    "body": {"contentType": "text", "content": "Bring the deck."},
    "attendees": [
        {"emailAddress": {"name": "Ana", "address": "ana@example.com"}, "type": "required"}
    ],
}


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def writer() -> GraphCalendarWriter:
    return GraphCalendarWriter(GraphClient(FakeTokenProvider()), SAO_PAULO)


def test_satisfies_the_calendar_writer_port(writer: GraphCalendarWriter) -> None:
    port: CalendarWriter = writer

    assert port is writer


@respx.mock
def test_posts_the_event_payload_to_the_default_calendar(writer: GraphCalendarWriter) -> None:
    route = respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json=CREATED))

    writer.create(A_NEW_EVENT)

    body = json.loads(route.calls.last.request.content)
    assert body["subject"] == "Planning"
    assert body["start"] == {"dateTime": "2026-09-15T09:00:00", "timeZone": "America/Sao_Paulo"}
    assert body["location"] == {"displayName": "Room 1"}


@respx.mock
def test_asks_for_the_created_event_in_the_resolved_zone(writer: GraphCalendarWriter) -> None:
    """No body: the Prefer header is exactly what it was before body support existed."""
    route = respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json=CREATED))

    writer.create(A_NEW_EVENT)

    assert route.calls.last.request.headers["Prefer"] == 'outlook.timezone="America/Sao_Paulo"'


@respx.mock
def test_sends_the_body_as_plain_text(writer: GraphCalendarWriter) -> None:
    route = respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json=CREATED_WITH_BODY))

    writer.create(A_NEW_EVENT_WITH_BODY)

    body = json.loads(route.calls.last.request.content)
    assert body["body"] == {"contentType": "text", "content": "Bring the deck."}


@respx.mock
def test_asks_for_the_zone_and_a_text_body_together_when_a_body_is_sent(
    writer: GraphCalendarWriter,
) -> None:
    route = respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json=CREATED_WITH_BODY))

    writer.create(A_NEW_EVENT_WITH_BODY)

    assert route.calls.last.request.headers["Prefer"] == (
        'outlook.timezone="America/Sao_Paulo", outlook.body-content-type="text"'
    )


@respx.mock
def test_reads_the_created_bodys_text_back(writer: GraphCalendarWriter) -> None:
    respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json=CREATED_WITH_BODY))

    created = writer.create(A_NEW_EVENT_WITH_BODY)

    assert created.body == "Bring the deck."


@respx.mock
def test_a_created_event_with_no_body_reads_back_empty(writer: GraphCalendarWriter) -> None:
    respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json=CREATED))

    created = writer.create(A_NEW_EVENT)

    assert created.body == ""


@respx.mock
def test_returns_the_event_as_graph_stored_it(writer: GraphCalendarWriter) -> None:
    respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json=CREATED))

    created = writer.create(A_NEW_EVENT)

    assert created.id == "AAMkNEW"
    assert created.start == NINE
    assert created.organizer.address == "me@example.com"


@respx.mock
def test_raises_when_graph_rejects_the_event(writer: GraphCalendarWriter) -> None:
    respx.post(EVENTS_URL).mock(
        return_value=httpx.Response(400, json={"error": {"code": "ErrorInvalidPropertyRequest"}})
    )

    with pytest.raises(GraphRequestError, match="ErrorInvalidPropertyRequest"):
        writer.create(A_NEW_EVENT)


@respx.mock
def test_raises_when_the_created_event_is_malformed(writer: GraphCalendarWriter) -> None:
    respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json={"id": "AAMkNEW"}))

    with pytest.raises(GraphResponseError):
        writer.create(A_NEW_EVENT)


@respx.mock
def test_fetches_an_event_by_id_with_the_detail_fields_selected(
    writer: GraphCalendarWriter,
) -> None:
    route = respx.get(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))

    event = writer.get_by_id("AAMkEXISTING")

    assert route.calls.last.request.url.params["$select"] == (
        "id,subject,start,end,isAllDay,location,organizer,body,attendees,"
        "reminderMinutesBeforeStart,sensitivity,showAs"
    )
    assert event.body == "Bring the deck."
    assert event.attendees == (EmailAddress(address="ana@example.com", display_name="Ana"),)


@respx.mock
def test_fetches_reminder_sensitivity_and_show_as(writer: GraphCalendarWriter) -> None:
    existing_with_extras = {
        **EXISTING,
        "reminderMinutesBeforeStart": 30,
        "sensitivity": "private",
        "showAs": "tentative",
    }
    respx.get(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=existing_with_extras))

    event = writer.get_by_id("AAMkEXISTING")

    assert event.reminder_minutes_before_start == 30
    assert event.sensitivity is Sensitivity.PRIVATE
    assert event.show_as is ShowAs.TENTATIVE


@respx.mock
def test_asks_for_the_zone_and_a_text_body_together_when_fetching_by_id(
    writer: GraphCalendarWriter,
) -> None:
    """Unlike create, this is unconditional: the event may already carry a body regardless
    of what this call touches.
    """
    route = respx.get(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))

    writer.get_by_id("AAMkEXISTING")

    assert route.calls.last.request.headers["Prefer"] == (
        'outlook.timezone="America/Sao_Paulo", outlook.body-content-type="text"'
    )


@respx.mock
def test_raises_event_not_found_when_getting_an_unknown_id(writer: GraphCalendarWriter) -> None:
    respx.get(AN_EVENT_URL).mock(
        return_value=httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})
    )

    with pytest.raises(EventNotFoundError):
        writer.get_by_id("AAMkEXISTING")


@respx.mock
def test_patches_only_the_supplied_fields(writer: GraphCalendarWriter) -> None:
    respx.get(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    route = respx.patch(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    changes = EventChanges(event_id="AAMkEXISTING", subject="Replanning")

    writer.update(changes)

    body = json.loads(route.calls.last.request.content)
    assert body == {"subject": "Replanning"}


@respx.mock
def test_patches_reminder_sensitivity_and_show_as(writer: GraphCalendarWriter) -> None:
    respx.get(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    route = respx.patch(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    changes = EventChanges(
        event_id="AAMkEXISTING",
        reminder_minutes_before_start=30,
        sensitivity=Sensitivity.CONFIDENTIAL,
        show_as=ShowAs.OOF,
    )

    writer.update(changes)

    body = json.loads(route.calls.last.request.content)
    assert body == {
        "reminderMinutesBeforeStart": 30,
        "sensitivity": "confidential",
        "showAs": "oof",
    }


@respx.mock
def test_patching_a_time_field_fetches_the_current_all_day_flag_first(
    writer: GraphCalendarWriter,
) -> None:
    get_route = respx.get(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    patch_route = respx.patch(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    new_start = NINE + timedelta(hours=2)
    changes = EventChanges(event_id="AAMkEXISTING", start=new_start)

    writer.update(changes)

    assert get_route.called
    body = json.loads(patch_route.calls.last.request.content)
    assert body["start"] == {"dateTime": "2026-09-15T11:00:00", "timeZone": "America/Sao_Paulo"}


@respx.mock
def test_patching_only_non_time_fields_skips_the_extra_fetch(
    writer: GraphCalendarWriter,
) -> None:
    get_route = respx.get(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    respx.patch(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    changes = EventChanges(event_id="AAMkEXISTING", location="Room 2")

    writer.update(changes)

    assert not get_route.called


@respx.mock
def test_reads_the_patched_event_back(writer: GraphCalendarWriter) -> None:
    respx.patch(AN_EVENT_URL).mock(return_value=httpx.Response(200, json=EXISTING))
    changes = EventChanges(event_id="AAMkEXISTING", subject="Replanning")

    updated = writer.update(changes)

    assert updated.id == "AAMkEXISTING"
    assert updated.attendees == (EmailAddress(address="ana@example.com", display_name="Ana"),)


@respx.mock
def test_raises_event_not_found_when_patching_an_unknown_id(writer: GraphCalendarWriter) -> None:
    respx.patch(AN_EVENT_URL).mock(
        return_value=httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})
    )
    changes = EventChanges(event_id="AAMkEXISTING", subject="Replanning")

    with pytest.raises(EventNotFoundError):
        writer.update(changes)


@respx.mock
def test_deletes_the_event_at_its_url(writer: GraphCalendarWriter) -> None:
    route = respx.delete(AN_EVENT_URL).mock(return_value=httpx.Response(204))

    writer.delete("AAMkEXISTING")

    assert route.called


@respx.mock
def test_raises_event_not_found_when_deleting_an_unknown_id(writer: GraphCalendarWriter) -> None:
    respx.delete(AN_EVENT_URL).mock(
        return_value=httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})
    )

    with pytest.raises(EventNotFoundError):
        writer.delete("AAMkEXISTING")


@respx.mock
def test_raises_when_graph_rejects_the_delete(writer: GraphCalendarWriter) -> None:
    respx.delete(AN_EVENT_URL).mock(
        return_value=httpx.Response(400, json={"error": {"code": "ErrorInvalidRequest"}})
    )

    with pytest.raises(GraphRequestError, match="ErrorInvalidRequest"):
        writer.delete("AAMkEXISTING")
