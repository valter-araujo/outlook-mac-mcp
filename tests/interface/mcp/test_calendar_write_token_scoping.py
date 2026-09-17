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
from outlook_mac_mcp.interface.mcp.calendar_update_tools import (
    PREVIEW_EVENT_UPDATE_TOOL,
    UPDATE_EVENT_TOOL,
)
from outlook_mac_mcp.interface.mcp.calendar_write_tools import CREATE_EVENT_TOOL, PREVIEW_EVENT_TOOL
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
A_NEW_EVENT: dict[str, Any] = {
    "subject": "Kickoff",
    "start": "2026-09-15T09:00:00-03:00",
    "end": "2026-09-15T10:00:00-03:00",
}
APPLY_TOOLS = {CREATE_EVENT_TOOL, UPDATE_EVENT_TOOL, DELETE_EVENT_TOOL}


def server_with_writes() -> tuple[MCPServer, InMemoryCalendarWriter]:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    return build_server(calendar_write_use_cases(writer)), writer


async def call(server: MCPServer, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(tool, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def create_token(server: MCPServer) -> str:
    draft = await call(server, PREVIEW_EVENT_TOOL, A_NEW_EVENT)
    return _token(draft)


async def update_token(server: MCPServer) -> str:
    draft = await call(
        server, PREVIEW_EVENT_UPDATE_TOOL, {"event_id": CURRENT.id, "subject": "Replanning"}
    )
    return _token(draft)


async def deletion_token(server: MCPServer) -> str:
    draft = await call(server, PREVIEW_EVENT_DELETION_TOOL, {"event_id": CURRENT.id})
    return _token(draft)


def _token(draft: dict[str, Any]) -> str:
    token = draft["token"]
    assert isinstance(token, str)
    return token


TOKEN_FACTORIES = {
    CREATE_EVENT_TOOL: create_token,
    UPDATE_EVENT_TOOL: update_token,
    DELETE_EVENT_TOOL: deletion_token,
}


@pytest.mark.parametrize("issuing_tool", sorted(APPLY_TOOLS))
async def test_a_token_is_accepted_only_by_its_own_apply_tool(issuing_tool: str) -> None:
    server, _ = server_with_writes()
    token = await TOKEN_FACTORIES[issuing_tool](server)

    for other_tool in APPLY_TOOLS - {issuing_tool}:
        with pytest.raises(ToolError):
            await call(server, other_tool, {"token": token})


async def test_a_create_token_does_not_update_or_delete_anything() -> None:
    server, writer = server_with_writes()
    token = await create_token(server)

    with pytest.raises(ToolError):
        await call(server, UPDATE_EVENT_TOOL, {"token": token})
    with pytest.raises(ToolError):
        await call(server, DELETE_EVENT_TOOL, {"token": token})

    assert writer.updated == []
    assert writer.deleted == []


async def test_an_update_token_does_not_create_or_delete_anything() -> None:
    server, writer = server_with_writes()
    token = await update_token(server)

    with pytest.raises(ToolError):
        await call(server, CREATE_EVENT_TOOL, {"token": token})
    with pytest.raises(ToolError):
        await call(server, DELETE_EVENT_TOOL, {"token": token})

    assert writer.created == []
    assert writer.deleted == []


async def test_a_deletion_token_does_not_create_or_update_anything() -> None:
    server, writer = server_with_writes()
    token = await deletion_token(server)

    with pytest.raises(ToolError):
        await call(server, CREATE_EVENT_TOOL, {"token": token})
    with pytest.raises(ToolError):
        await call(server, UPDATE_EVENT_TOOL, {"token": token})

    assert writer.created == []
    assert writer.updated == []


async def test_each_tokens_own_apply_tool_still_works_after_the_others_reject_it() -> None:
    """Rejection by a foreign apply tool must not consume the token: it was never found
    in that tool's store, so nothing was taken from the store that does hold it.
    """
    server, writer = server_with_writes()
    token = await update_token(server)

    with pytest.raises(ToolError):
        await call(server, DELETE_EVENT_TOOL, {"token": token})

    updated = await call(server, UPDATE_EVENT_TOOL, {"token": token})

    assert updated["subject"] == "Replanning"
    assert len(writer.updated) == 1
