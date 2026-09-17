from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.interface.mcp.calendar_deletion_tools import (
    DELETE_EVENT_TOOL,
    PREVIEW_EVENT_DELETION_TOOL,
)
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_calendar_writer import InMemoryCalendarWriter
from tests.fakes.use_case_bundles import calendar_write_use_cases

pytestmark = pytest.mark.anyio

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
ORGANIZER = EmailAddress(address="me@example.com", display_name="Me")
CURRENT = Event(
    id="AAMkEXISTING",
    subject="Planning",
    start=NINE,
    end=NINE + timedelta(hours=1),
    is_all_day=False,
    location="Room 1",
    organizer=ORGANIZER,
    body="Bring the deck.",
    attendees=(EmailAddress(address="ana@example.com"),),
)


def server_with_writes() -> tuple[MCPServer, InMemoryCalendarWriter]:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    return build_server(calendar_write_use_cases(writer)), writer


async def tool_names(server: MCPServer) -> set[str]:
    return {tool.name for tool in await server.list_tools()}


async def call(server: MCPServer, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(tool, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def description_of(server: MCPServer, name: str) -> str:
    return next(tool for tool in await server.list_tools() if tool.name == name).description or ""


async def test_registers_the_deletion_pair_when_the_flag_is_on() -> None:
    server, _ = server_with_writes()

    assert {PREVIEW_EVENT_DELETION_TOOL, DELETE_EVENT_TOOL} <= await tool_names(server)


async def test_preview_shows_the_full_current_event_without_deleting() -> None:
    server, writer = server_with_writes()

    draft = await call(server, PREVIEW_EVENT_DELETION_TOOL, {"event_id": CURRENT.id})

    assert draft["token"]
    assert 'subject: "Planning"' in draft["summary"]
    assert 'location: "Room 1"' in draft["summary"]
    assert 'body: "Bring the deck."' in draft["summary"]
    assert "attendees: ana@example.com" in draft["summary"]
    assert writer.deleted == []


async def test_delete_with_the_token_removes_exactly_the_previewed_event() -> None:
    server, writer = server_with_writes()
    token = (await call(server, PREVIEW_EVENT_DELETION_TOOL, {"event_id": CURRENT.id}))["token"]

    result = await call(server, DELETE_EVENT_TOOL, {"token": token})

    assert result["event_id"] == CURRENT.id
    assert result["deleted"] is True
    assert writer.deleted == [CURRENT.id]


async def test_a_token_works_exactly_once() -> None:
    server, writer = server_with_writes()
    token = (await call(server, PREVIEW_EVENT_DELETION_TOOL, {"event_id": CURRENT.id}))["token"]
    await call(server, DELETE_EVENT_TOOL, {"token": token})

    with pytest.raises(ToolError, match="preview again"):
        await call(server, DELETE_EVENT_TOOL, {"token": token})

    assert len(writer.deleted) == 1


async def test_refuses_a_token_it_never_issued() -> None:
    server, writer = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, DELETE_EVENT_TOOL, {"token": "made-up"})

    assert writer.deleted == []


async def test_refuses_an_unknown_event_id_at_preview() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_DELETION_TOOL, {"event_id": "never-seeded"})


async def test_delete_takes_only_the_token() -> None:
    server, _ = server_with_writes()
    tool = next(tool for tool in await server.list_tools() if tool.name == DELETE_EVENT_TOOL)

    assert set(tool.input_schema["properties"]) == {"token"}


async def test_delete_says_it_is_irreversible_and_needs_a_same_session_token() -> None:
    server, _ = server_with_writes()

    description = await description_of(server, DELETE_EVENT_TOOL)

    assert "IRREVERSIBLE" in description
    assert "preview_event_deletion" in description
    assert "same session" in description
    assert "exactly once" in description


async def test_preview_describes_showing_the_full_event_for_disambiguation() -> None:
    server, _ = server_with_writes()

    description = await description_of(server, PREVIEW_EVENT_DELETION_TOOL)

    assert "unabbreviated" in description


@pytest.mark.parametrize("tool", [PREVIEW_EVENT_DELETION_TOOL, DELETE_EVENT_TOOL])
async def test_both_tools_require_confirmation_of_details_taken_from_email(tool: str) -> None:
    server, _ = server_with_writes()

    description = await description_of(server, tool)

    assert "email content" in description
    assert "confirmation" in description
    assert "BEFORE calling preview_event_deletion" in description
