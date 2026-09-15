import json
from datetime import UTC, datetime, timedelta
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.errors import NotAuthenticatedError
from outlook_mac_mcp.interface.mcp.mail_listing_tools import (
    COUNT_EMAILS_TOOL,
    LIST_EMAILS_TOOL,
)
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.use_case_bundles import mail_only_use_cases
from tests.interface.mcp.test_server import FailingMailRepository

pytestmark = pytest.mark.anyio

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)
FILTER_ARGUMENTS = {
    "folder",
    "is_read",
    "sender",
    "received_after",
    "received_before",
    "has_attachments",
}


def make_email(
    email_id: str,
    *,
    days_ago: int = 0,
    is_read: bool = False,
    sender: str = "ana@example.com",
    has_attachments: bool = False,
) -> Email:
    return Email(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address=sender),
        received_at=BASE_TIME - timedelta(days=days_ago),
        is_read=is_read,
        has_attachments=has_attachments,
        preview="",
    )


def server_with(*emails: Email) -> MCPServer:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(FolderName.INBOX, email)
    return build_server(mail_only_use_cases(repository))


async def call(server: MCPServer, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(tool, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def listed_ids(server: MCPServer, arguments: dict[str, Any]) -> list[str]:
    return [item["id"] for item in (await call(server, LIST_EMAILS_TOOL, arguments))["items"]]


async def tool_named(server: MCPServer, name: str) -> Any:
    return next(tool for tool in await server.list_tools() if tool.name == name)


async def test_list_offers_every_filter_plus_sort_and_limit() -> None:
    tool = await tool_named(server_with(), LIST_EMAILS_TOOL)

    assert set(tool.input_schema["properties"]) == FILTER_ARGUMENTS | {"sort", "limit"}
    assert tool.input_schema["properties"]["limit"]["minimum"] == MIN_LIMIT
    assert tool.input_schema["properties"]["limit"]["maximum"] == MAX_LIMIT


async def test_count_offers_the_filters_only() -> None:
    tool = await tool_named(server_with(), COUNT_EMAILS_TOOL)

    assert set(tool.input_schema["properties"]) == FILTER_ARGUMENTS


async def test_offers_exactly_the_two_sort_orders() -> None:
    tool = await tool_named(server_with(), LIST_EMAILS_TOOL)

    assert set(tool.input_schema["$defs"]["SortOrder"]["enum"]) == {"newest", "oldest"}


async def test_lists_newest_first_by_default_with_exact_totals() -> None:
    server = server_with(
        make_email("old", days_ago=2), make_email("new"), make_email("mid", days_ago=1)
    )

    page = await call(server, LIST_EMAILS_TOOL, {"limit": 2})

    assert [item["id"] for item in page["items"]] == ["new", "mid"]
    assert page["returned"] == 2
    assert page["total"] == 3
    assert page["total_is_exact"] is True


async def test_lists_oldest_first_when_asked() -> None:
    server = server_with(make_email("old", days_ago=2), make_email("new"))

    assert await listed_ids(server, {"sort": "oldest"}) == ["old", "new"]


async def test_passes_every_filter_through() -> None:
    server = server_with(
        make_email("hit", days_ago=1, is_read=True, sender="bo@example.com", has_attachments=True),
        make_email("other-sender", days_ago=1, is_read=True, has_attachments=True),
        make_email(
            "too-old", days_ago=9, is_read=True, sender="bo@example.com", has_attachments=True
        ),
        make_email("unread", days_ago=1, sender="bo@example.com", has_attachments=True),
        make_email("no-attachment", days_ago=1, is_read=True, sender="bo@example.com"),
    )

    listed = await listed_ids(
        server,
        {
            "is_read": True,
            "sender": "bo@example.com",
            "received_after": (BASE_TIME - timedelta(days=3)).isoformat(),
            "received_before": BASE_TIME.isoformat(),
            "has_attachments": True,
        },
    )

    assert listed == ["hit"]


async def test_returns_an_empty_exact_page_when_nothing_matches() -> None:
    page = await call(server_with(), LIST_EMAILS_TOOL, {"sender": "nobody@example.com"})

    assert page["items"] == []
    assert page["total"] == 0
    assert page["total_is_exact"] is True


async def test_counts_everything_by_default() -> None:
    server = server_with(make_email("a"), make_email("b"), make_email("c"))

    assert await call(server, COUNT_EMAILS_TOOL, {}) == {"total": 3}


async def test_counts_only_what_the_filters_admit() -> None:
    server = server_with(make_email("read", is_read=True), make_email("unread"))

    assert await call(server, COUNT_EMAILS_TOOL, {"is_read": False}) == {"total": 1}


@pytest.mark.parametrize("tool", [LIST_EMAILS_TOOL, COUNT_EMAILS_TOOL])
async def test_rejects_a_sender_that_is_not_an_address(tool: str) -> None:
    with pytest.raises(ToolError):
        await call(server_with(), tool, {"sender": "not an address"})


@pytest.mark.parametrize("tool", [LIST_EMAILS_TOOL, COUNT_EMAILS_TOOL])
async def test_rejects_a_date_bound_without_an_offset(tool: str) -> None:
    with pytest.raises(ToolError):
        await call(server_with(), tool, {"received_after": "2026-09-01T00:00:00"})


async def test_rejects_an_inverted_date_range() -> None:
    with pytest.raises(ToolError):
        await call(
            server_with(),
            LIST_EMAILS_TOOL,
            {"received_after": BASE_TIME.isoformat(), "received_before": BASE_TIME.isoformat()},
        )


async def test_rejects_an_unknown_sort() -> None:
    with pytest.raises(ToolError):
        await call(server_with(), LIST_EMAILS_TOOL, {"sort": "random"})


@pytest.mark.parametrize("tool", [LIST_EMAILS_TOOL, COUNT_EMAILS_TOOL])
async def test_translates_a_project_error_into_a_tool_error(tool: str) -> None:
    repository = FailingMailRepository(NotAuthenticatedError("run the sign-in command first"))
    server = build_server(mail_only_use_cases(repository))

    with pytest.raises(ToolError, match="run the sign-in command first"):
        await call(server, tool, {})


async def test_list_says_it_can_answer_how_many_and_how_far_back() -> None:
    description = (await tool_named(server_with(), LIST_EMAILS_TOOL)).description or ""

    assert "sort=oldest" in description
    assert "exact" in description
    assert '"showing N of M"' in description


async def test_count_says_what_it_cannot_tell() -> None:
    description = (await tool_named(server_with(), COUNT_EMAILS_TOOL)).description or ""

    assert "exact total" in description
    assert "cannot tell which emails" in description


async def test_logs_the_count_as_the_item_count(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()
    server = server_with(make_email("a"), make_email("b"))

    await call(server, COUNT_EMAILS_TOOL, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == COUNT_EMAILS_TOOL
    assert record["item_count"] == 2
