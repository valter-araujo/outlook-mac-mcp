from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.domain.new_event import MAX_ATTENDEES, MAX_BODY_LENGTH, MAX_SUBJECT_LENGTH
from outlook_mac_mcp.domain.show_as import ShowAs
from outlook_mac_mcp.interface.mcp.calendar_write_tools import (
    CREATE_EVENT_TOOL,
    PREVIEW_EVENT_TOOL,
)
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_calendar_writer import InMemoryCalendarWriter
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.use_case_bundles import calendar_write_use_cases, mail_only_use_cases

pytestmark = pytest.mark.anyio

WRITE_TOOLS = {PREVIEW_EVENT_TOOL, CREATE_EVENT_TOOL}
A_PREVIEW: dict[str, Any] = {
    "subject": "Planning",
    "start": "2026-09-15T09:00:00-03:00",
    "end": "2026-09-15T10:00:00-03:00",
    "location": "Room 1",
    "attendees": ["ana@example.com"],
}


def server_with_writes() -> tuple[MCPServer, InMemoryCalendarWriter]:
    writer = InMemoryCalendarWriter()
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


async def test_registers_no_write_tools_when_the_flag_is_off() -> None:
    server = build_server(mail_only_use_cases(InMemoryMailRepository()))

    assert not (await tool_names(server)) & WRITE_TOOLS


async def test_registers_both_write_tools_when_the_flag_is_on() -> None:
    server, _ = server_with_writes()

    assert await tool_names(server) >= WRITE_TOOLS


async def test_preview_returns_a_token_and_a_summary_without_creating() -> None:
    server, writer = server_with_writes()

    draft = await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW)

    assert draft["token"]
    assert "Planning" in draft["summary"]
    assert "ana@example.com" in draft["summary"]
    assert writer.created == []


async def test_create_with_the_token_creates_exactly_the_previewed_event() -> None:
    server, writer = server_with_writes()
    token = (await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW))["token"]

    created = await call(server, CREATE_EVENT_TOOL, {"token": token})

    assert created["id"] == "created-1"
    assert created["subject"] == "Planning"
    assert created["start"] == "2026-09-15T09:00:00-03:00"
    assert created["body"] == ""
    assert [attendee.address for attendee in writer.created[0].attendees] == ["ana@example.com"]


async def test_omitting_the_body_preserves_the_previous_behaviour_exactly() -> None:
    """A_PREVIEW carries no "body" key at all: the same shape of call this project
    always accepted, before body support existed.
    """
    assert "body" not in A_PREVIEW
    server, writer = server_with_writes()

    draft = await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW)
    created = await call(server, CREATE_EVENT_TOOL, {"token": draft["token"]})

    assert "body" not in draft["summary"]
    assert writer.created[0].body == ""
    assert created["body"] == ""


async def test_preview_includes_a_short_body_in_full_in_the_summary() -> None:
    server, _ = server_with_writes()

    draft = await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"body": "Bring the deck."})

    assert 'body: "Bring the deck."' in draft["summary"]


async def test_create_carries_the_body_through_to_the_created_event() -> None:
    server, writer = server_with_writes()
    token = (await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"body": "Bring the deck."}))[
        "token"
    ]

    created = await call(server, CREATE_EVENT_TOOL, {"token": token})

    assert created["body"] == "Bring the deck."
    assert writer.created[0].body == "Bring the deck."


async def test_rejects_a_body_beyond_the_maximum_length() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"body": "a" * (MAX_BODY_LENGTH + 1)})


async def test_preview_includes_reminder_sensitivity_and_show_as_in_the_summary() -> None:
    server, _ = server_with_writes()

    draft = await call(
        server,
        PREVIEW_EVENT_TOOL,
        A_PREVIEW
        | {"reminder_minutes_before_start": 30, "sensitivity": "private", "show_as": "tentative"},
    )

    assert "reminder: 30 min before" in draft["summary"]
    assert "sensitivity: private" in draft["summary"]
    assert "show as: tentative" in draft["summary"]


async def test_create_carries_reminder_sensitivity_and_show_as_through() -> None:
    server, writer = server_with_writes()
    token = (
        await call(
            server,
            PREVIEW_EVENT_TOOL,
            A_PREVIEW | {"reminder_minutes_before_start": 30, "show_as": "busy"},
        )
    )["token"]

    await call(server, CREATE_EVENT_TOOL, {"token": token})

    assert writer.created[0].reminder_minutes_before_start == 30
    assert writer.created[0].show_as is ShowAs.BUSY


async def test_rejects_a_negative_reminder() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"reminder_minutes_before_start": -1})


async def test_rejects_an_unknown_sensitivity() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"sensitivity": "top-secret"})


async def test_rejects_an_unknown_show_as() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"show_as": "somewhere"})


async def test_a_token_works_exactly_once() -> None:
    server, writer = server_with_writes()
    token = (await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW))["token"]
    await call(server, CREATE_EVENT_TOOL, {"token": token})

    with pytest.raises(ToolError, match="preview again"):
        await call(server, CREATE_EVENT_TOOL, {"token": token})

    assert len(writer.created) == 1


async def test_refuses_a_token_it_never_issued() -> None:
    server, writer = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, CREATE_EVENT_TOOL, {"token": "made-up"})

    assert writer.created == []


async def test_rejects_a_start_without_an_offset() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"start": "2026-09-15T09:00:00"})


async def test_rejects_an_end_before_the_start() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"end": "2026-09-15T08:00:00-03:00"})


async def test_rejects_a_subject_beyond_the_maximum_length() -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(
            server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"subject": "a" * (MAX_SUBJECT_LENGTH + 1)}
        )


async def test_rejects_more_attendees_than_allowed() -> None:
    server, _ = server_with_writes()
    too_many = [f"p{index}@example.com" for index in range(MAX_ATTENDEES + 1)]

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"attendees": too_many})


@pytest.mark.parametrize(
    "address", ["not-an-address", "two@@example.com", "spaced name@example.com"]
)
async def test_rejects_an_attendee_that_is_not_an_address(address: str) -> None:
    server, _ = server_with_writes()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_EVENT_TOOL, A_PREVIEW | {"attendees": [address]})


async def test_create_takes_only_the_token() -> None:
    server, _ = server_with_writes()
    tool = next(tool for tool in await server.list_tools() if tool.name == CREATE_EVENT_TOOL)

    assert set(tool.input_schema["properties"]) == {"token"}


async def test_create_says_it_is_a_real_change_needing_a_same_session_token() -> None:
    server, _ = server_with_writes()

    description = await description_of(server, CREATE_EVENT_TOOL)

    assert "REAL CHANGE" in description
    assert "preview_event" in description
    assert "same session" in description
    assert "exactly once" in description


@pytest.mark.parametrize("tool", [PREVIEW_EVENT_TOOL, CREATE_EVENT_TOOL])
async def test_both_tools_require_confirmation_of_details_taken_from_email(tool: str) -> None:
    server, _ = server_with_writes()

    description = await description_of(server, tool)

    assert "email content" in description
    assert "confirmation" in description
    assert "BEFORE calling preview_event" in description


@pytest.mark.parametrize("tool", [PREVIEW_EVENT_TOOL, CREATE_EVENT_TOOL])
async def test_both_tools_also_require_confirmation_of_body_content_taken_from_the_web(
    tool: str,
) -> None:
    server, _ = server_with_writes()

    description = await description_of(server, tool)

    assert "web content" in description
    assert "the body" in description
    assert "BEFORE calling preview_event" in description
