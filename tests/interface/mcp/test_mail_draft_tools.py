from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.domain.new_mail_draft import MAX_BODY_LENGTH, MAX_RECIPIENTS
from outlook_mac_mcp.interface.mcp.calendar_write_tools import CREATE_EVENT_TOOL
from outlook_mac_mcp.interface.mcp.mail_draft_tools import CREATE_DRAFT_TOOL, PREVIEW_DRAFT_TOOL
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_calendar_writer import InMemoryCalendarWriter
from tests.fakes.in_memory_mail_draft_writer import InMemoryMailDraftWriter
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.use_case_bundles import (
    calendar_write_use_cases,
    mail_draft_use_cases,
    mail_only_use_cases,
)

pytestmark = pytest.mark.anyio

DRAFT_TOOLS = {PREVIEW_DRAFT_TOOL, CREATE_DRAFT_TOOL}
A_PREVIEW: dict[str, Any] = {
    "subject": "Hello",
    "body": "Hi Ana.",
    "to": ["ana@example.com"],
    "cc": ["bo@example.com"],
}


def server_with_drafts() -> tuple[MCPServer, InMemoryMailDraftWriter]:
    writer = InMemoryMailDraftWriter()
    return build_server(mail_draft_use_cases(writer)), writer


async def tool_names(server: MCPServer) -> set[str]:
    return {tool.name for tool in await server.list_tools()}


async def call(server: MCPServer, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(tool, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def test_registers_no_draft_tools_when_the_flag_is_off() -> None:
    server = build_server(mail_only_use_cases(InMemoryMailRepository()))

    assert not (await tool_names(server)) & DRAFT_TOOLS


async def test_registers_no_draft_tools_when_only_calendar_write_is_on() -> None:
    server = build_server(calendar_write_use_cases(InMemoryCalendarWriter()))

    assert not (await tool_names(server)) & DRAFT_TOOLS


async def test_registers_both_draft_tools_when_the_flag_is_on() -> None:
    server, _ = server_with_drafts()

    assert await tool_names(server) >= DRAFT_TOOLS


async def test_registers_no_calendar_write_tools_when_only_the_draft_flag_is_on() -> None:
    server, _ = server_with_drafts()

    assert CREATE_EVENT_TOOL not in await tool_names(server)


async def test_preview_returns_a_token_and_a_summary_without_saving() -> None:
    server, writer = server_with_drafts()

    preview = await call(server, PREVIEW_DRAFT_TOOL, A_PREVIEW)

    assert preview["token"]
    assert "ana@example.com" in preview["summary"]
    assert 'body: "Hi Ana."' in preview["summary"]
    assert writer.created == []


async def test_create_saves_exactly_the_previewed_draft() -> None:
    server, writer = server_with_drafts()
    token = (await call(server, PREVIEW_DRAFT_TOOL, A_PREVIEW))["token"]

    created = await call(server, CREATE_DRAFT_TOOL, {"token": token})

    assert created["id"] == "draft-1"
    assert [recipient.address for recipient in writer.created[0].to] == ["ana@example.com"]
    assert [recipient.address for recipient in writer.created[0].cc] == ["bo@example.com"]


async def test_create_states_that_the_draft_was_not_sent() -> None:
    server, _ = server_with_drafts()
    token = (await call(server, PREVIEW_DRAFT_TOOL, A_PREVIEW))["token"]

    created = await call(server, CREATE_DRAFT_TOOL, {"token": token})

    assert "NOT been sent" in created["notice"]


async def test_a_token_saves_at_most_one_draft() -> None:
    server, writer = server_with_drafts()
    token = (await call(server, PREVIEW_DRAFT_TOOL, A_PREVIEW))["token"]
    await call(server, CREATE_DRAFT_TOOL, {"token": token})

    with pytest.raises(ToolError):
        await call(server, CREATE_DRAFT_TOOL, {"token": token})

    assert len(writer.created) == 1


async def test_refuses_a_token_that_was_never_issued() -> None:
    server, writer = server_with_drafts()

    with pytest.raises(ToolError, match="preview again"):
        await call(server, CREATE_DRAFT_TOOL, {"token": "never-issued"})

    assert writer.created == []


@pytest.mark.parametrize(
    "overrides",
    [
        {"to": []},
        {"to": ["not-an-address"]},
        {"cc": ["no at sign.example.com"]},
        {"subject": ""},
        {"body": "a" * (MAX_BODY_LENGTH + 1)},
        {"to": [f"p{index}@example.com" for index in range(MAX_RECIPIENTS + 1)]},
    ],
)
async def test_rejects_an_invalid_preview_before_anything_is_parked(
    overrides: dict[str, Any],
) -> None:
    server, writer = server_with_drafts()

    with pytest.raises(ToolError):
        await call(server, PREVIEW_DRAFT_TOOL, A_PREVIEW | overrides)

    assert writer.created == []


async def test_describes_both_tools_with_the_untrusted_content_rule() -> None:
    server, _ = server_with_drafts()

    descriptions = {tool.name: tool.description or "" for tool in await server.list_tools()}

    assert "untrusted" in descriptions[PREVIEW_DRAFT_TOOL]
    assert "untrusted" in descriptions[CREATE_DRAFT_TOOL]
    assert "NOT sent" in descriptions[CREATE_DRAFT_TOOL]
