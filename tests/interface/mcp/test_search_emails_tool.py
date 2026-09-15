from datetime import UTC, datetime, timedelta
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.application.search_emails_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.interface.mcp.server import SEARCH_EMAILS_TOOL, build_server
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

pytestmark = pytest.mark.anyio

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)
OPERATOR_LIKE_TERMS = [
    "from:ceo@example.com",
    "deck AND from:ana@example.com",
    'deck" AND from:ceo@example.com "',
    "subject:payroll OR NOT lunch",
    "ratio 3:1",
]


def make_email(email_id: str, *, subject: str, minutes_ago: int = 0) -> Email:
    return Email(
        id=email_id,
        subject=subject,
        sender=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
        received_at=BASE_TIME - timedelta(minutes=minutes_ago),
        is_read=False,
        has_attachments=False,
        preview="",
    )


def server_with(*emails: Email) -> MCPServer:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(FolderName.INBOX, email)
    return build_server(
        ListUnreadEmails(repository), SearchEmails(repository), GetEmail(repository)
    )


async def call_search(server: MCPServer, arguments: dict[str, Any]) -> list[dict[str, Any]]:
    result = await server.call_tool(SEARCH_EMAILS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    assert result.structured_content is not None
    items = result.structured_content["result"]
    assert isinstance(items, list)
    return items


async def tool_schema(server: MCPServer) -> dict[str, Any]:
    tools = await server.list_tools()
    tool = next(tool for tool in tools if tool.name == SEARCH_EMAILS_TOOL)
    properties = tool.input_schema["properties"]
    assert isinstance(properties, dict)
    return properties


async def test_registers_the_tool_with_term_folder_scope_and_limit() -> None:
    assert set(await tool_schema(server_with())) == {"term", "folder", "scope", "limit"}


async def test_advertises_the_term_length_bounds() -> None:
    term = (await tool_schema(server_with()))["term"]

    assert term["minLength"] == MIN_TERM_LENGTH
    assert term["maxLength"] == MAX_TERM_LENGTH


async def test_advertises_the_limit_bounds() -> None:
    limit = (await tool_schema(server_with()))["limit"]

    assert limit["minimum"] == MIN_LIMIT
    assert limit["maximum"] == MAX_LIMIT


async def test_says_in_its_description_that_results_are_ranked_by_relevance() -> None:
    tools = await server_with().list_tools()

    description = next(t for t in tools if t.name == SEARCH_EMAILS_TOOL).description or ""
    assert "relevance" in description.lower()
    assert "not by date" in description.lower()


async def test_returns_the_matching_emails() -> None:
    server = server_with(
        make_email("hit", subject="Quarterly review"), make_email("miss", subject="Lunch")
    )

    items = await call_search(server, {"term": "quarterly"})

    assert [item["id"] for item in items] == ["hit"]


async def test_searches_the_requested_folder() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("inboxed", subject="deck"))
    repository.add(FolderName.ARCHIVE, make_email("archived", subject="deck"))
    server = build_server(
        ListUnreadEmails(repository), SearchEmails(repository), GetEmail(repository)
    )

    items = await call_search(server, {"term": "deck", "folder": "archive"})

    assert [item["id"] for item in items] == ["archived"]


async def test_returns_an_empty_list_when_nothing_matches() -> None:
    assert await call_search(server_with(), {"term": "payroll"}) == []


@pytest.mark.parametrize("term", OPERATOR_LIKE_TERMS)
async def test_accepts_a_term_that_reads_like_a_query(term: str) -> None:
    """These must reach the search as ordinary text, not be rejected or interpreted."""
    assert await call_search(server_with(), {"term": term}) == []


async def test_finds_an_email_whose_subject_contains_operator_like_text() -> None:
    server = server_with(make_email("hit", subject="Re: from:ana AND the deck"))

    items = await call_search(server, {"term": "from:ana AND the deck"})

    assert [item["id"] for item in items] == ["hit"]


async def test_requires_a_term() -> None:
    with pytest.raises(ToolError):
        await call_search(server_with(), {})


async def test_rejects_an_empty_term() -> None:
    with pytest.raises(ToolError):
        await call_search(server_with(), {"term": ""})


async def test_rejects_a_term_beyond_the_maximum_length() -> None:
    with pytest.raises(ToolError):
        await call_search(server_with(), {"term": "a" * (MAX_TERM_LENGTH + 1)})


async def test_rejects_a_limit_outside_the_advertised_range() -> None:
    with pytest.raises(ToolError):
        await call_search(server_with(), {"term": "deck", "limit": MAX_LIMIT + 1})


async def test_rejects_an_unknown_folder() -> None:
    with pytest.raises(ToolError):
        await call_search(server_with(), {"term": "deck", "folder": "nowhere"})


async def test_offers_exactly_the_three_scopes_and_nothing_else() -> None:
    tools = await server_with().list_tools()
    tool = next(t for t in tools if t.name == SEARCH_EMAILS_TOOL)

    scopes = tool.input_schema["$defs"]["SearchScope"]["enum"]

    assert set(scopes) == {"any", "subject", "sender"}


async def test_defaults_the_scope_to_any() -> None:
    assert (await tool_schema(server_with()))["scope"]["default"] == SearchScope.ANY


async def test_a_subject_scope_matches_only_the_subject() -> None:
    server = server_with(
        make_email("in-subject", subject="Contoso application"),
        make_email("in-preview", subject="Job digest"),
    )

    items = await call_search(server, {"term": "Contoso", "scope": "subject"})

    assert [item["id"] for item in items] == ["in-subject"]


async def test_a_sender_scope_matches_only_the_sender() -> None:
    repository = InMemoryMailRepository()
    repository.add(
        FolderName.INBOX,
        Email(
            id="from-contoso",
            subject="Application received",
            sender=EmailAddress(address="donotreply@contoso.com"),
            received_at=BASE_TIME,
            is_read=False,
            has_attachments=False,
            preview="",
        ),
    )
    repository.add(FolderName.INBOX, make_email("about-contoso", subject="Contoso is hiring"))
    server = build_server(
        ListUnreadEmails(repository), SearchEmails(repository), GetEmail(repository)
    )

    items = await call_search(server, {"term": "contoso", "scope": "sender"})

    assert [item["id"] for item in items] == ["from-contoso"]


async def test_the_any_scope_matches_subject_and_sender_alike() -> None:
    server = server_with(make_email("hit", subject="Contoso is hiring"))

    items = await call_search(server, {"term": "Contoso", "scope": "any"})

    assert [item["id"] for item in items] == ["hit"]


async def test_rejects_an_unknown_scope() -> None:
    with pytest.raises(ToolError):
        await call_search(server_with(), {"term": "Contoso", "scope": "body"})


async def test_says_in_its_description_that_results_are_capped_without_a_total() -> None:
    tools = await server_with().list_tools()

    description = (next(t for t in tools if t.name == SEARCH_EMAILS_TOOL).description or "").lower()
    assert "at most" in description
    assert "no total match count" in description
    assert "seen them all" in description
