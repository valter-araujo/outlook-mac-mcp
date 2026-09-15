import json
from datetime import UTC, datetime, timedelta
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError, UnexpectedToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.list_unread_emails import (
    MAX_LIMIT,
    MIN_LIMIT,
    ListUnreadEmails,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.errors import NotAuthenticatedError
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.server import LIST_UNREAD_EMAILS_TOOL, build_server
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

pytestmark = pytest.mark.anyio

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)
SECRET_MESSAGE = "mailbox of ceo@example.com is unreachable"


class FailingMailRepository:
    """Raises on every read, to drive the error-translation paths."""

    def __init__(self, error: Exception) -> None:
        self._error = error

    def list_unread(self, folder: FolderName, limit: int) -> tuple[Email, ...]:
        raise self._error


def make_email(email_id: str, *, minutes_ago: int = 0) -> Email:
    return Email(
        id=email_id,
        subject="Quarterly review",
        sender=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
        received_at=BASE_TIME - timedelta(minutes=minutes_ago),
        is_read=False,
        has_attachments=False,
        preview="preview",
    )


def server_with(*emails: Email) -> MCPServer:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(FolderName.INBOX, email)
    return build_server(ListUnreadEmails(repository))


def server_failing_with(error: Exception) -> MCPServer:
    return build_server(ListUnreadEmails(FailingMailRepository(error)))


async def call_tool(server: MCPServer, arguments: dict[str, Any]) -> list[dict[str, Any]]:
    result = await server.call_tool(LIST_UNREAD_EMAILS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    assert result.structured_content is not None
    items = result.structured_content["result"]
    assert isinstance(items, list)
    return items


async def tool_properties(server: MCPServer) -> dict[str, Any]:
    tools = await server.list_tools()
    tool = next(tool for tool in tools if tool.name == LIST_UNREAD_EMAILS_TOOL)
    properties = tool.input_schema["properties"]
    assert isinstance(properties, dict)
    return properties


async def test_registers_the_tool_with_flat_arguments() -> None:
    properties = await tool_properties(server_with())

    assert set(properties) == {"folder", "limit"}


async def test_advertises_the_limit_bounds_in_the_schema() -> None:
    properties = await tool_properties(server_with())

    assert properties["limit"]["minimum"] == MIN_LIMIT
    assert properties["limit"]["maximum"] == MAX_LIMIT


async def test_returns_the_unread_emails_newest_first() -> None:
    server = server_with(make_email("newest"), make_email("older", minutes_ago=10))

    items = await call_tool(server, {})

    assert [item["id"] for item in items] == ["newest", "older"]


async def test_projects_the_sender_into_flat_fields() -> None:
    items = await call_tool(server_with(make_email("one")), {})

    assert items[0]["sender_address"] == "ana@example.com"
    assert items[0]["sender_name"] == "Ana Lima"


async def test_serializes_the_timestamp_as_iso_8601() -> None:
    items = await call_tool(server_with(make_email("one")), {})

    assert items[0]["received_at"] == "2026-09-14T12:00:00Z"


async def test_passes_the_folder_and_limit_through_to_the_use_case() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.ARCHIVE, make_email("archived"))
    repository.add(FolderName.INBOX, make_email("inboxed"))
    server = build_server(ListUnreadEmails(repository))

    items = await call_tool(server, {"folder": "archive", "limit": 1})

    assert [item["id"] for item in items] == ["archived"]


async def test_returns_an_empty_list_when_nothing_is_unread() -> None:
    assert await call_tool(server_with(), {}) == []


async def test_translates_a_project_error_into_a_tool_error_carrying_its_message() -> None:
    server = server_failing_with(NotAuthenticatedError("run the sign-in command first"))

    with pytest.raises(ToolError, match="run the sign-in command first"):
        await call_tool(server, {})


async def test_does_not_translate_an_unexpected_error() -> None:
    server = server_failing_with(ZeroDivisionError(SECRET_MESSAGE))

    with pytest.raises(UnexpectedToolError):
        await call_tool(server, {})


async def test_rejects_a_limit_outside_the_advertised_range() -> None:
    with pytest.raises(ToolError):
        await call_tool(server_with(), {"limit": MAX_LIMIT + 1})


async def test_rejects_an_unknown_folder() -> None:
    with pytest.raises(ToolError):
        await call_tool(server_with(), {"folder": "nowhere"})


async def test_logs_the_call_with_its_item_count_and_never_touches_stdout(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()
    server = server_with(make_email("one"), make_email("two", minutes_ago=5))

    await call_tool(server, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == LIST_UNREAD_EMAILS_TOOL
    assert record["succeeded"] is True
    assert record["item_count"] == 2


async def test_logs_a_failure_by_error_type_without_its_message(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()
    server = server_failing_with(NotAuthenticatedError(SECRET_MESSAGE))

    with pytest.raises(ToolError):
        await call_tool(server, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["succeeded"] is False
    assert record["error_type"] == "NotAuthenticatedError"
    assert "ceo@example.com" not in captured.err
