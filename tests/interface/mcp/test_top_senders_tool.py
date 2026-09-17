import json
from datetime import UTC, datetime, timedelta
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.top_senders import DEFAULT_TOP_SENDERS, SCAN_CEILING
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.errors import NotAuthenticatedError
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.server import build_server
from outlook_mac_mcp.interface.mcp.top_senders_tool import TOP_SENDERS_TOOL
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.use_case_bundles import mail_only_use_cases
from tests.interface.mcp.test_server import FailingMailRepository

pytestmark = pytest.mark.anyio

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)


def make_email(email_id: str, sender: str, *, name: str = "", days_ago: int = 0) -> Email:
    return Email(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address=sender, display_name=name),
        received_at=BASE_TIME - timedelta(days=days_ago),
        is_read=False,
        has_attachments=False,
        preview="",
    )


def server_with(*emails: Email) -> MCPServer:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(FolderName.INBOX, email)
    return build_server(mail_only_use_cases(repository))


async def call(server: MCPServer, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(TOP_SENDERS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def the_tool(server: MCPServer) -> Any:
    return next(tool for tool in await server.list_tools() if tool.name == TOP_SENDERS_TOOL)


async def test_offers_folder_read_state_date_range_and_limit() -> None:
    tool = await the_tool(server_with())

    properties = tool.input_schema["properties"]
    assert set(properties) == {"folder", "is_read", "received_after", "received_before", "limit"}
    assert properties["limit"]["default"] == DEFAULT_TOP_SENDERS
    assert properties["limit"]["minimum"] == MIN_LIMIT
    assert properties["limit"]["maximum"] == MAX_LIMIT


async def test_ranks_senders_with_their_names_and_counts() -> None:
    server = server_with(
        make_email("1", "bo@x.io", name="Bo"),
        make_email("2", "ana@x.io", name="Ana Lima"),
        make_email("3", "bo@x.io"),
    )

    ranking = await call(server, {})

    assert ranking["senders"] == [
        {"address": "bo@x.io", "name": "Bo", "count": 2},
        {"address": "ana@x.io", "name": "Ana Lima", "count": 1},
    ]
    assert ranking["scanned"] == 3
    assert ranking["total"] == 3
    assert ranking["coverage_is_complete"] is True


async def test_folder_all_merges_every_well_known_folder() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("1", "bo@x.io", name="Bo"))
    repository.add(FolderName.ARCHIVE, make_email("2", "bo@x.io", name="Bo"))
    server = build_server(mail_only_use_cases(repository))

    ranking = await call(server, {"folder": "all"})

    assert ranking["senders"] == [{"address": "bo@x.io", "name": "Bo", "count": 2}]
    assert ranking["total"] == 2
    assert ranking["folder"] == "inbox, archive, junkemail, sentitems, drafts"


async def test_applies_the_filters_and_the_limit() -> None:
    server = server_with(
        make_email("1", "old@x.io", days_ago=40),
        make_email("2", "a@x.io"),
        make_email("3", "b@x.io"),
        make_email("4", "b@x.io"),
    )

    ranking = await call(
        server, {"received_after": (BASE_TIME - timedelta(days=7)).isoformat(), "limit": 1}
    )

    assert [item["address"] for item in ranking["senders"]] == ["b@x.io"]
    assert ranking["total"] == 3


async def test_ranks_nobody_on_an_empty_folder() -> None:
    ranking = await call(server_with(), {})

    assert ranking == {
        "senders": [],
        "scanned": 0,
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


async def test_says_the_ranking_covers_only_scanned_emails_when_incomplete() -> None:
    description = (await the_tool(server_with())).description or ""

    assert f"{SCAN_CEILING:,}" in description
    assert "ONLY the scanned emails" in description
    assert "received_after" in description
    assert "coverage_is_complete" in description


async def test_logs_the_number_of_ranked_senders_and_no_address(
    capsys: pytest.CaptureFixture[str],
) -> None:
    configure_logging()
    server = server_with(make_email("1", "ceo@x.io"), make_email("2", "cfo@x.io"))

    await call(server, {})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == TOP_SENDERS_TOOL
    assert record["item_count"] == 2
    assert "x.io" not in captured.err
