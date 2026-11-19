from datetime import UTC, datetime
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.domain.saved_draft import SavedDraft
from outlook_mac_mcp.interface.mcp.draft_deletion_tools import (
    DELETE_DRAFT_TOOL,
    PREVIEW_DRAFT_DELETION_TOOL,
)
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_draft_deleter import InMemoryDraftDeleter
from tests.fakes.in_memory_mail_draft_writer import InMemoryMailDraftWriter
from tests.fakes.use_case_bundles import mail_draft_use_cases

pytestmark = pytest.mark.anyio

DELETION_TOOLS = {PREVIEW_DRAFT_DELETION_TOOL, DELETE_DRAFT_TOOL}
A_DRAFT = SavedDraft(
    id="AAMkDRAFT", subject="Hello", to=(), last_modified_at=datetime(2026, 11, 17, tzinfo=UTC)
)


def server_with_deletion() -> tuple[MCPServer, InMemoryDraftDeleter]:
    deleter = InMemoryDraftDeleter(drafts=(A_DRAFT,), other_ids=frozenset({"AAMkINBOX"}))
    return build_server(mail_draft_use_cases(InMemoryMailDraftWriter(), deleter=deleter)), deleter


async def tool_names(server: MCPServer) -> set[str]:
    return {tool.name for tool in await server.list_tools()}


async def call(server: MCPServer, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(tool, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def test_registers_no_deletion_tools_when_only_the_draft_flag_is_on() -> None:
    server = build_server(mail_draft_use_cases(InMemoryMailDraftWriter()))

    assert not (await tool_names(server)) & DELETION_TOOLS


async def test_registers_both_deletion_tools_when_the_delete_flag_is_on() -> None:
    server, _ = server_with_deletion()

    assert await tool_names(server) >= DELETION_TOOLS


async def test_deletes_the_previewed_draft_and_says_where_it_went() -> None:
    server, deleter = server_with_deletion()
    preview = await call(server, PREVIEW_DRAFT_DELETION_TOOL, {"email_id": "AAMkDRAFT"})

    result = await call(server, DELETE_DRAFT_TOOL, {"token": preview["token"]})

    assert deleter.deleted == ["AAMkDRAFT"]
    assert "Deleted Items" in result["notice"]


async def test_preview_refuses_a_message_outside_drafts() -> None:
    server, _ = server_with_deletion()

    with pytest.raises(ToolError, match="not a draft"):
        await server.call_tool(PREVIEW_DRAFT_DELETION_TOOL, {"email_id": "AAMkINBOX"})
