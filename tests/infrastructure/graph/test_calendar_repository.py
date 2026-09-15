from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.calendar_repository import CalendarRepository
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.time_window import TimeWindow
from outlook_mac_mcp.infrastructure.graph.calendar_repository import (
    MAX_PAGES,
    PAGE_SIZE,
    GraphCalendarRepository,
)
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError

CALENDAR_VIEW_URL = f"{GRAPH_BASE_URL}/me/calendarView"
SAO_PAULO = ZoneInfo("America/Sao_Paulo")
MIDNIGHT = datetime(2026, 9, 15, tzinfo=SAO_PAULO)
TODAY = TimeWindow(start=MIDNIGHT, end=MIDNIGHT + timedelta(days=1))
NOTHING: dict[str, Any] = {"value": []}


def graph_event(event_id: str, *, start: str = "2026-09-15T09:00:00.0000000") -> dict[str, Any]:
    return {
        "id": event_id,
        "subject": "Planning",
        "start": {"dateTime": start, "timeZone": "America/Sao_Paulo"},
        "end": {"dateTime": "2026-09-15T10:00:00.0000000", "timeZone": "America/Sao_Paulo"},
        "isAllDay": False,
        "location": {"displayName": "Room 1"},
        "organizer": {"emailAddress": {"name": "Ana Lima", "address": "ana@example.com"}},
    }


def page_link(number: int) -> str:
    return f"{CALENDAR_VIEW_URL}?%24skip={number * PAGE_SIZE}"


def linked_page(event_ids: list[str], next_page: int | None) -> dict[str, Any]:
    page: dict[str, Any] = {"value": [graph_event(event_id) for event_id in event_ids]}
    if next_page is not None:
        page["@odata.nextLink"] = page_link(next_page)
    return page


def mock_page(number: int, payload: dict[str, Any]) -> respx.Route:
    """Followed pages are told apart from the first request by their $skip."""
    return respx.get(CALENDAR_VIEW_URL, params__contains={"$skip": str(number * PAGE_SIZE)}).mock(
        return_value=httpx.Response(200, json=payload)
    )


def mock_first_page(payload: dict[str, Any]) -> respx.Route:
    return respx.get(CALENDAR_VIEW_URL).mock(return_value=httpx.Response(200, json=payload))


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def repository() -> GraphCalendarRepository:
    return GraphCalendarRepository(GraphClient(FakeTokenProvider()), SAO_PAULO)


def test_satisfies_the_calendar_repository_port(repository: GraphCalendarRepository) -> None:
    port: CalendarRepository = repository

    assert port is repository


@respx.mock
def test_sends_the_window_as_iso_8601_with_offsets(repository: GraphCalendarRepository) -> None:
    route = mock_first_page(NOTHING)

    repository.list_events(TODAY)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["startDateTime"] == "2026-09-15T00:00:00-03:00"
    assert parameters["endDateTime"] == "2026-09-16T00:00:00-03:00"


@respx.mock
def test_asks_for_times_in_the_resolved_zone(repository: GraphCalendarRepository) -> None:
    route = mock_first_page(NOTHING)

    repository.list_events(TODAY)

    assert route.calls.last.request.headers["Prefer"] == 'outlook.timezone="America/Sao_Paulo"'


@respx.mock
def test_selects_only_the_fields_the_entity_needs_ordered_by_start(
    repository: GraphCalendarRepository,
) -> None:
    route = mock_first_page(NOTHING)

    repository.list_events(TODAY)

    parameters = dict(route.calls.last.request.url.params)
    assert set(parameters["$select"].split(",")) == {
        "id",
        "subject",
        "start",
        "end",
        "isAllDay",
        "location",
        "organizer",
    }
    assert parameters["$orderby"] == "start/dateTime"
    assert parameters["$top"] == str(PAGE_SIZE)


@respx.mock
def test_maps_the_events_in_the_order_graph_gave_them(
    repository: GraphCalendarRepository,
) -> None:
    mock_first_page({"value": [graph_event("first"), graph_event("second")]})

    page = repository.list_events(TODAY)

    assert [event.id for event in page.items] == ["first", "second"]
    assert page.items[0].start == datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)


@respx.mock
def test_returns_an_empty_exact_page_when_the_window_is_free(
    repository: GraphCalendarRepository,
) -> None:
    mock_first_page(NOTHING)

    assert repository.list_events(TODAY) == Page(items=(), total=0, total_is_exact=True)


@respx.mock
def test_follows_every_page_and_reports_an_exact_total(
    repository: GraphCalendarRepository,
) -> None:
    mock_page(1, linked_page(["c"], next_page=2))
    last = mock_page(2, linked_page(["d"], next_page=None))
    mock_first_page(linked_page(["a", "b"], next_page=1))

    page = repository.list_events(TODAY)

    assert [event.id for event in page.items] == ["a", "b", "c", "d"]
    assert page.total == 4
    assert page.total_is_exact is True
    assert last.call_count == 1


@respx.mock
def test_stops_after_the_maximum_number_of_pages_and_reports_a_lower_bound(
    repository: GraphCalendarRepository,
) -> None:
    for number in range(1, MAX_PAGES + 1):
        mock_page(number, linked_page([f"page-{number}"], next_page=number + 1))
    beyond = mock_page(MAX_PAGES + 1, linked_page(["never"], next_page=None))
    mock_first_page(linked_page(["page-0"], next_page=1))

    page = repository.list_events(TODAY)

    assert len(page.items) == MAX_PAGES
    assert page.total == MAX_PAGES
    assert page.total_is_exact is False
    assert beyond.call_count == 0


@respx.mock
def test_raises_when_a_page_has_no_value_array(repository: GraphCalendarRepository) -> None:
    mock_first_page({"error": None})

    with pytest.raises(GraphResponseError):
        repository.list_events(TODAY)


@respx.mock
def test_raises_when_a_followed_page_is_malformed(repository: GraphCalendarRepository) -> None:
    mock_page(1, {"value": ["not-an-event"]})
    mock_first_page(linked_page(["a"], next_page=1))

    with pytest.raises(GraphResponseError):
        repository.list_events(TODAY)
