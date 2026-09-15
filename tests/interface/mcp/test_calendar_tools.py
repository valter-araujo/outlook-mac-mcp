import json
from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.list_upcoming_events import (
    DEFAULT_UPCOMING_DAYS,
    MAX_UPCOMING_DAYS,
    MIN_UPCOMING_DAYS,
)
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.time_window import TimeWindow
from outlook_mac_mcp.infrastructure.graph.errors import NotAuthenticatedError
from outlook_mac_mcp.interface.mcp.calendar_tools import (
    LIST_TODAYS_EVENTS_TOOL,
    LIST_UPCOMING_EVENTS_TOOL,
)
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.fixed_clock import FixedClock
from tests.fakes.in_memory_calendar_repository import InMemoryCalendarRepository
from tests.fakes.use_case_bundles import calendar_only_use_cases

pytestmark = pytest.mark.anyio

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
MIDNIGHT = datetime(2026, 9, 15, tzinfo=SAO_PAULO)
NOW = MIDNIGHT + timedelta(hours=15, minutes=42)


class FailingCalendarRepository:
    def __init__(self, error: Exception) -> None:
        self._error = error

    def list_events(self, window: TimeWindow) -> Page[Event]:
        raise self._error


def make_event(event_id: str, start: datetime, *, is_all_day: bool = False) -> Event:
    return Event(
        id=event_id,
        subject="Planning",
        start=start,
        end=start + timedelta(days=1 if is_all_day else 0, hours=0 if is_all_day else 1),
        is_all_day=is_all_day,
        location="Room 1",
        organizer=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
    )


def server_with(*events: Event) -> MCPServer:
    repository = InMemoryCalendarRepository()
    for event in events:
        repository.add(event)
    return build_server(calendar_only_use_cases(repository, FixedClock(NOW)))


async def call_page(server: MCPServer, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(tool, arguments)
    assert isinstance(result, CallToolResult)
    page = result.structured_content
    assert isinstance(page, dict)
    return page


async def tool_named(server: MCPServer, name: str) -> Any:
    tools = await server.list_tools()
    return next(tool for tool in tools if tool.name == name)


async def test_todays_tool_takes_no_arguments() -> None:
    tool = await tool_named(server_with(), LIST_TODAYS_EVENTS_TOOL)

    assert tool.input_schema.get("properties", {}) == {}


async def test_upcoming_tool_advertises_the_days_bounds_and_default() -> None:
    tool = await tool_named(server_with(), LIST_UPCOMING_EVENTS_TOOL)

    days = tool.input_schema["properties"]["days"]
    assert days["minimum"] == MIN_UPCOMING_DAYS
    assert days["maximum"] == MAX_UPCOMING_DAYS
    assert days["default"] == DEFAULT_UPCOMING_DAYS


async def test_returns_todays_events_as_flat_records_earliest_first() -> None:
    server = server_with(
        make_event("later", MIDNIGHT + timedelta(hours=14)),
        make_event("earlier", MIDNIGHT + timedelta(hours=9)),
        make_event("tomorrow", MIDNIGHT + timedelta(hours=33)),
    )

    page = await call_page(server, LIST_TODAYS_EVENTS_TOOL, {})

    assert [item["id"] for item in page["items"]] == ["earlier", "later"]
    assert page["items"][0]["start"] == "2026-09-15T09:00:00-03:00"
    assert page["items"][0]["organizer_address"] == "ana@example.com"
    assert page["items"][0]["organizer_name"] == "Ana Lima"
    assert page["items"][0]["location"] == "Room 1"


async def test_reports_returned_and_total_for_today() -> None:
    server = server_with(make_event("one", MIDNIGHT + timedelta(hours=9)))

    page = await call_page(server, LIST_TODAYS_EVENTS_TOOL, {})

    assert page["returned"] == 1
    assert page["total"] == 1
    assert page["total_is_exact"] is True


async def test_reports_zero_of_zero_on_a_free_day() -> None:
    page = await call_page(server_with(), LIST_TODAYS_EVENTS_TOOL, {})

    assert page["items"] == []
    assert page["returned"] == 0
    assert page["total"] == 0
    assert page["total_is_exact"] is True


async def test_includes_an_all_day_event_flagged_as_such() -> None:
    server = server_with(make_event("all-day", MIDNIGHT, is_all_day=True))

    page = await call_page(server, LIST_TODAYS_EVENTS_TOOL, {})

    assert page["items"][0]["is_all_day"] is True


async def test_returns_the_upcoming_events_within_the_requested_days() -> None:
    server = server_with(
        make_event("day-two", NOW + timedelta(days=2)),
        make_event("day-five", NOW + timedelta(days=5)),
    )

    page = await call_page(server, LIST_UPCOMING_EVENTS_TOOL, {"days": 3})

    assert [item["id"] for item in page["items"]] == ["day-two"]
    assert page["total"] == 1


async def test_upcoming_defaults_to_a_week() -> None:
    server = server_with(
        make_event("day-six", NOW + timedelta(days=6)),
        make_event("day-eight", NOW + timedelta(days=8)),
    )

    page = await call_page(server, LIST_UPCOMING_EVENTS_TOOL, {})

    assert [item["id"] for item in page["items"]] == ["day-six"]


@pytest.mark.parametrize("days", [MIN_UPCOMING_DAYS - 1, MAX_UPCOMING_DAYS + 1])
async def test_rejects_days_outside_the_advertised_range(days: int) -> None:
    with pytest.raises(ToolError):
        await call_page(server_with(), LIST_UPCOMING_EVENTS_TOOL, {"days": days})


@pytest.mark.parametrize("tool", [LIST_TODAYS_EVENTS_TOOL, LIST_UPCOMING_EVENTS_TOOL])
async def test_tells_the_client_to_say_showing_n_of_m(tool: str) -> None:
    description = (await tool_named(server_with(), tool)).description or ""

    assert '"showing N of M"' in description
    assert "at least" in description


@pytest.mark.parametrize("tool", [LIST_TODAYS_EVENTS_TOOL, LIST_UPCOMING_EVENTS_TOOL])
async def test_translates_a_project_error_into_a_tool_error(tool: str) -> None:
    repository = FailingCalendarRepository(NotAuthenticatedError("run the sign-in command first"))
    server = build_server(calendar_only_use_cases(repository, FixedClock(NOW)))

    with pytest.raises(ToolError, match="run the sign-in command first"):
        await call_page(server, tool, {})


async def test_logs_the_call_with_its_item_count(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()
    server = server_with(
        make_event("one", MIDNIGHT + timedelta(hours=9)),
        make_event("two", MIDNIGHT + timedelta(hours=10)),
    )

    await call_page(server, LIST_TODAYS_EVENTS_TOOL, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == LIST_TODAYS_EVENTS_TOOL
    assert record["item_count"] == 2
