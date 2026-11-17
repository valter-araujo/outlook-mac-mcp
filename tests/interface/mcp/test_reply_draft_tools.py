from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.interface.mcp.reply_draft_tools import (
    CREATE_REPLY_DRAFT_TOOL,
    PREVIEW_REPLY_DRAFT_TOOL,
)
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.application.reply_fixtures import ME, NO_REPLY_MESSAGE, THREAD_MESSAGE
from tests.fakes.in_memory_mail_draft_writer import InMemoryMailDraftWriter
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.in_memory_reply_draft_writer import InMemoryReplyDraftWriter
from tests.fakes.use_case_bundles import mail_draft_use_cases, mail_only_use_cases

pytestmark = pytest.mark.anyio

REPLY_TOOLS = {PREVIEW_REPLY_DRAFT_TOOL, CREATE_REPLY_DRAFT_TOOL}


def server_with_replies() -> tuple[MCPServer, InMemoryReplyDraftWriter]:
    replies = InMemoryReplyDraftWriter(
        sources=(THREAD_MESSAGE, NO_REPLY_MESSAGE), own_addresses=frozenset({ME})
    )
    return build_server(mail_draft_use_cases(InMemoryMailDraftWriter(), replies)), replies


async def tool_names(server: MCPServer) -> set[str]:
    return {tool.name for tool in await server.list_tools()}


async def call(server: MCPServer, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(tool, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def test_registers_no_reply_tools_when_the_draft_flag_is_off() -> None:
    server = build_server(mail_only_use_cases(InMemoryMailRepository()))

    assert not (await tool_names(server)) & REPLY_TOOLS


async def test_registers_both_reply_tools_when_the_draft_flag_is_on() -> None:
    server, _ = server_with_replies()

    assert await tool_names(server) >= REPLY_TOOLS


async def test_preview_defaults_to_reply_all_and_returns_the_final_lines() -> None:
    server, replies = server_with_replies()

    preview = await call(server, PREVIEW_REPLY_DRAFT_TOOL, {"email_id": "AAMk1", "body": "Ok"})

    assert preview["to"] == ["ana@example.com", "bo@example.com"]
    assert preview["cc"] == ["cy@example.com", "di@example.com"]
    assert replies.created == []


async def test_preview_returns_the_no_reply_warning() -> None:
    server, _ = server_with_replies()

    preview = await call(server, PREVIEW_REPLY_DRAFT_TOOL, {"email_id": "AAMk2", "body": "Ok"})

    assert len(preview["warnings"]) == 1


async def test_create_saves_the_previewed_reply_and_says_it_was_not_sent() -> None:
    server, replies = server_with_replies()
    preview = await call(
        server, PREVIEW_REPLY_DRAFT_TOOL, {"email_id": "AAMk1", "body": "Ok", "mode": "reply"}
    )

    created = await call(server, CREATE_REPLY_DRAFT_TOOL, {"token": preview["token"]})

    assert created["id"] == "reply-1"
    assert "NOT been sent" in created["notice"]
    assert [r.address for r in replies.created[0].recipients.to] == ["ana@example.com"]


async def test_create_refuses_a_reused_token() -> None:
    server, _ = server_with_replies()
    preview = await call(server, PREVIEW_REPLY_DRAFT_TOOL, {"email_id": "AAMk1", "body": "Ok"})
    await call(server, CREATE_REPLY_DRAFT_TOOL, {"token": preview["token"]})

    with pytest.raises(ToolError, match="preview again"):
        await server.call_tool(CREATE_REPLY_DRAFT_TOOL, {"token": preview["token"]})


async def test_preview_rejects_an_unknown_mode() -> None:
    server, _ = server_with_replies()

    with pytest.raises(ToolError):
        await server.call_tool(
            PREVIEW_REPLY_DRAFT_TOOL, {"email_id": "AAMk1", "body": "Ok", "mode": "forward"}
        )


async def test_preview_rejects_an_invalid_extra_address() -> None:
    server, _ = server_with_replies()

    with pytest.raises(ToolError):
        await server.call_tool(
            PREVIEW_REPLY_DRAFT_TOOL,
            {"email_id": "AAMk1", "body": "Ok", "extra_cc": ["not an address"]},
        )


async def test_a_new_draft_token_is_refused_by_create_reply_draft() -> None:
    server, _ = server_with_replies()
    new_draft = await call(
        server, "preview_draft", {"subject": "Hi", "body": "Ok", "to": ["ana@example.com"]}
    )

    with pytest.raises(ToolError, match="preview again"):
        await server.call_tool(CREATE_REPLY_DRAFT_TOOL, {"token": new_draft["token"]})
