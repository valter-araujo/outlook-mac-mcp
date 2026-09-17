from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.interface.mcp.calendar_update_tools import (
    PREVIEW_EVENT_UPDATE_TOOL,
    UPDATE_EVENT_TOOL,
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


async def test_registers_the_update_pair_when_the_flag_is_on() -> None:
    server, _ = server_with_writes()

    assert {PREVIEW_EVENT_UPDATE_TOOL, UPDATE_EVENT_TOOL} <= await tool_names(server)


async def test_preview_shows_the_diff_without_updating() -> None:
    server, writer = server_with_writes()

    draft = await call(
        server, PREVIEW_EVENT_UPDATE_TOOL, {"event_id": CURRENT.id, "subject": "Replanning"}
    )

    assert draft["token"]
    assert 'subject: "Planning" -> "Replanning"' in draft["summary"]
    assert writer.updated == []


async def test_update_with_the_token_applies_exactly_the_previewed_change() -> None:
    server, writer = server_with_writes()
    token = (
        await call(
            server, PREVIEW_EVENT_UPDATE_TOOL, {"event_id": CURRENT.id, "subject": "Replanning"}
        )
    )["token"]

    updated = await call(server, UPDATE_EVENT_TOOL, {"token": token})

    assert updated["subject"] == "Replanning"
    assert len(writer.updated) == 1
    assert writer.updated[0].subject == "Replanning"


async def test_a_field_left_out_of_the_call_is_not_changed() -> None:
    server, writer = server_with_writes()
    token = (
        await call(
            server, PREVIEW_EVENT_UPDATE_TOOL, {"event_id": CURRENT.id, "location": "Room 2"}
        )
    )["token"]

    updated = await call(server, UPDATE_EVENT_TOOL, {"token": token})

    assert updated["subject"] == "Planning"
    assert writer.updated[0].subject is None


async def test_a_token_works_exactly_once() -> None:
    server, writer = server_with_writes()
    token = (
        await call(
            server, PREVIEW_EVENT_UPDATE_TOOL, {"event_id": CURRENT.id, "subject": "Replanning"}
        )
    )["token"]
    await call(server, UPDATE_EVENT_TOOL, {"token": token})

    with pytest.raises(ToolError, match="preview again"):
        await call(server, UPDATE_EVENT_TOOL, {"token": token})

    assert len(writer.updated) == 1


async def test_refuses_a_token_it_never_issued() -> None:
    server, writer = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, UPDATE_EVENT_TOOL, {"token": "made-up"})

    assert writer.updated == []


async def test_refuses_an_unknown_event_id_at_preview() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(
            server,
            PREVIEW_EVENT_UPDATE_TOOL,
            {"event_id": "never-seeded", "subject": "Replanning"},
        )


async def test_preview_takes_no_fields_beyond_event_id_and_the_updatable_ones() -> None:
    server, _ = server_with_writes()
    tool = next(
        tool for tool in await server.list_tools() if tool.name == PREVIEW_EVENT_UPDATE_TOOL
    )

    assert set(tool.input_schema["properties"]) == {
        "event_id",
        "subject",
        "start",
        "end",
        "location",
        "body",
        "attendees",
    }


async def test_update_says_it_is_a_real_change_needing_a_same_session_token() -> None:
    server, _ = server_with_writes()

    description = await description_of(server, UPDATE_EVENT_TOOL)

    assert "REAL CHANGE" in description
    assert "preview_event_update" in description
    assert "same session" in description
    assert "exactly once" in description


async def test_preview_describes_that_only_supplied_fields_change() -> None:
    server, _ = server_with_writes()

    description = await description_of(server, PREVIEW_EVENT_UPDATE_TOOL)

    assert "left out" in description or "left exactly as" in description


@pytest.mark.parametrize("tool", [PREVIEW_EVENT_UPDATE_TOOL, UPDATE_EVENT_TOOL])
async def test_both_tools_require_confirmation_of_details_taken_from_email(tool: str) -> None:
    server, _ = server_with_writes()

    description = await description_of(server, tool)

    assert "email content" in description
    assert "confirmation" in description
    assert "BEFORE calling preview_event_update" in description
