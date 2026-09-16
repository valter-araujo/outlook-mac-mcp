import json
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.infrastructure.graph.calendar_writer import GraphCalendarWriter
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError

EVENTS_URL = f"{GRAPH_BASE_URL}/me/events"
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
