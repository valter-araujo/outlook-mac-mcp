import json
from datetime import UTC, datetime, timedelta
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_largest_emails import (
    DEFAULT_LARGEST_EMAILS,
    SCAN_CEILING,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.errors import NotAuthenticatedError
from outlook_mac_mcp.interface.mcp.list_largest_emails_tool import LIST_LARGEST_EMAILS_TOOL
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.use_case_bundles import mail_only_use_cases
from tests.interface.mcp.test_server import FailingMailRepository

pytestmark = pytest.mark.anyio

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)


def make_email(email_id: str, *, days_ago: int = 0, is_read: bool = False) -> Email:
    return Email(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
        received_at=BASE_TIME - timedelta(days=days_ago),
        is_read=is_read,
        has_attachments=False,
        preview="",
    )


def server_with(*sized: tuple[Email, int | None]) -> MCPServer:
    """A size of None stands in for a message with no size property at all."""
    repository = InMemoryMailRepository()
    for email, size_bytes in sized:
        repository.add(FolderName.INBOX, email, size_bytes=size_bytes)
    return build_server(mail_only_use_cases(repository))


async def call(server: MCPServer, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(LIST_LARGEST_EMAILS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def the_tool(server: MCPServer) -> Any:
    return next(tool for tool in await server.list_tools() if tool.name == LIST_LARGEST_EMAILS_TOOL)


async def test_offers_folder_read_state_date_range_and_limit() -> None:
    tool = await the_tool(server_with())

    properties = tool.input_schema["properties"]
    assert set(properties) == {"folder", "is_read", "received_after", "limit"}
    assert properties["limit"]["default"] == DEFAULT_LARGEST_EMAILS
    assert properties["limit"]["minimum"] == MIN_LIMIT
    assert properties["limit"]["maximum"] == MAX_LIMIT


async def test_ranks_emails_largest_first_with_their_details() -> None:
    server = server_with(
        (make_email("small"), 100),
        (make_email("huge", days_ago=1), 90_000),
    )

    ranking = await call(server, {})

    assert [item["id"] for item in ranking["items"]] == ["huge", "small"]
    huge = ranking["items"][0]
    assert huge["subject"] == "subject"
    assert huge["sender_address"] == "ana@example.com"
    assert huge["sender_name"] == "Ana Lima"
    assert huge["size_bytes"] == 90_000
    assert ranking["scanned"] == 2
    assert ranking["skipped"] == 0
    assert ranking["total"] == 2
    assert ranking["coverage_is_complete"] is True


async def test_an_email_with_no_known_size_is_skipped_not_ranked() -> None:
    server = server_with((make_email("sized"), 100), (make_email("unsized", days_ago=1), None))

    ranking = await call(server, {})

    assert [item["id"] for item in ranking["items"]] == ["sized"]
    assert ranking["skipped"] == 1
    assert ranking["scanned"] == 2
    assert ranking["total"] == 2


async def test_folder_all_merges_every_well_known_folder() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("inboxed"), size_bytes=10)
    repository.add(FolderName.ARCHIVE, make_email("archived", days_ago=1), size_bytes=20)
    server = build_server(mail_only_use_cases(repository))

    ranking = await call(server, {"folder": "all"})

    assert [item["id"] for item in ranking["items"]] == ["archived", "inboxed"]
    assert ranking["total"] == 2
    assert ranking["folder"] == "inbox, archive, junkemail, sentitems, drafts"


async def test_applies_the_filters_and_the_limit() -> None:
    server = server_with(
        (make_email("old", days_ago=40), 999_999),
        (make_email("small"), 100),
        (make_email("mid", days_ago=1), 500),
        (make_email("large", days_ago=2), 9_000),
    )

    ranking = await call(
        server, {"received_after": (BASE_TIME - timedelta(days=7)).isoformat(), "limit": 1}
    )

    assert [item["id"] for item in ranking["items"]] == ["large"]
    assert ranking["total"] == 3


async def test_ranks_nothing_on_an_empty_folder() -> None:
    ranking = await call(server_with(), {})

    assert ranking == {
        "items": [],
        "scanned": 0,
        "skipped": 0,
        "total": 0,
        "coverage_is_complete": True,
        "folder": "inbox",
    }


async def test_rejects_a_date_bound_without_an_offset() -> None:
    with pytest.raises(ToolError):
        await call(server_with(), {"received_after": "2026-09-01T00:00:00"})


async def test_rejects_a_limit_outside_the_advertised_range() -> None:
    with pytest.raises(ToolError):
        await call(server_with(), {"limit": MAX_LIMIT + 1})


async def test_translates_a_project_error_into_a_tool_error() -> None:
    repository = FailingMailRepository(NotAuthenticatedError("run the sign-in command first"))
    server = build_server(mail_only_use_cases(repository))

    with pytest.raises(ToolError, match="run the sign-in command first"):
        await call(server, {})


async def test_says_size_comes_from_an_extended_property_and_coverage_can_be_partial() -> None:
    description = (await the_tool(server_with())).description or ""

    assert "extended property" in description
    assert f"{SCAN_CEILING:,}" in description
    assert "ONLY the scanned emails" in description
    assert "received_after" in description
    assert "coverage_is_complete" in description
    assert "skipped" in description


async def test_logs_the_number_of_ranked_emails_and_no_subject_or_size(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()
    server = server_with((make_email("1"), 424242), (make_email("2", days_ago=1), 999999))

    await call(server, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == LIST_LARGEST_EMAILS_TOOL
    assert record["item_count"] == 2
    assert "424242" not in captured.err
    assert "999999" not in captured.err
