from datetime import UTC, datetime, timedelta
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.search_emails_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
    SearchEmailsRequest,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.email_size_scan import EmailSizeScan
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.domain.sender_scan import SenderScan
from outlook_mac_mcp.interface.mcp.server import SEARCH_EMAILS_TOOL, build_server
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository
from tests.fakes.use_case_bundles import mail_only_use_cases

pytestmark = pytest.mark.anyio

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)
OPERATOR_LIKE_TERMS = [
    "from:ceo@example.com",
    "deck AND from:ana@example.com",
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


A_LOWER_BOUND = 250


class LowerBoundMailRepository:
    """A backend that stopped counting, the way Graph does past the search ceiling."""

    def list_unread(self, folder: FolderName, limit: int) -> Page[Email]:
        raise NotImplementedError

    def get_by_id(self, email_id: str) -> EmailDetail:
        raise NotImplementedError

    def search(self, request: SearchEmailsRequest) -> Page[Email]:
        return Page(
            items=(make_email("hit", subject="deck"),), total=A_LOWER_BOUND, total_is_exact=False
        )

    def list_matching(self, request: ListEmailsRequest) -> Page[Email]:
        raise NotImplementedError

    def count_matching(self, filters: EmailFilters) -> int:
        raise NotImplementedError

    def scan_senders(self, filters: EmailFilters, ceiling: int) -> SenderScan:
        raise NotImplementedError

    def scan_email_sizes(self, filters: EmailFilters, ceiling: int) -> EmailSizeScan:
        raise NotImplementedError


def server_with(*emails: Email) -> MCPServer:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(FolderName.INBOX, email)
    return build_server(mail_only_use_cases(repository))


async def search_page(server: MCPServer, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(SEARCH_EMAILS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    page = result.structured_content
    assert isinstance(page, dict)
    return page


async def call_search(server: MCPServer, arguments: dict[str, Any]) -> list[dict[str, Any]]:
    items = (await search_page(server, arguments))["items"]
    assert isinstance(items, list)
    return items


async def tool_description(server: MCPServer) -> str:
    tools = await server.list_tools()
    return next(tool for tool in tools if tool.name == SEARCH_EMAILS_TOOL).description or ""


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
    description = await tool_description(server_with())

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
    server = build_server(mail_only_use_cases(repository))

    items = await call_search(server, {"term": "deck", "folder": "archive"})

    assert [item["id"] for item in items] == ["archived"]


async def test_returns_an_empty_list_when_nothing_matches() -> None:
    assert await call_search(server_with(), {"term": "payroll"}) == []


async def test_reports_zero_totals_when_nothing_matches() -> None:
    page = await search_page(server_with(), {"term": "payroll"})

    assert page["returned"] == 0
    assert page["total"] == 0
    assert page["total_is_exact"] is True


async def test_reports_how_many_were_returned_out_of_how_many_match() -> None:
    server = server_with(*(make_email(str(index), subject="deck") for index in range(5)))

    page = await search_page(server, {"term": "deck", "limit": 2})

    assert page["returned"] == 2
    assert page["total"] == 5
    assert page["total_is_exact"] is True


async def test_passes_an_inexact_total_through_as_a_lower_bound() -> None:
    repository = LowerBoundMailRepository()
    server = build_server(mail_only_use_cases(repository))

    page = await search_page(server, {"term": "deck"})

    assert page["returned"] == 1
    assert page["total"] == A_LOWER_BOUND
    assert page["total_is_exact"] is False


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
    server = build_server(mail_only_use_cases(repository))

    items = await call_search(server, {"term": "contoso", "scope": "sender"})

    assert [item["id"] for item in items] == ["from-contoso"]


async def test_the_any_scope_matches_subject_and_sender_alike() -> None:
    server = server_with(make_email("hit", subject="Contoso is hiring"))

    items = await call_search(server, {"term": "Contoso", "scope": "any"})

    assert [item["id"] for item in items] == ["hit"]


async def test_rejects_an_unknown_scope() -> None:
    with pytest.raises(ToolError):
        await call_search(server_with(), {"term": "Contoso", "scope": "body"})


async def test_says_results_are_not_a_sample_to_infer_totals_or_dates_from() -> None:
    description = await tool_description(server_with())

    assert "NOT a representative sample" in description
    assert "do not infer a total, a date range" in description
    assert "count_emails" in description
    assert "list_emails with sort=oldest" in description


async def test_tells_the_client_to_say_showing_n_of_m_and_to_narrow_the_scope() -> None:
    description = await tool_description(server_with())

    assert '"showing N of M"' in description
    assert "at least" in description
    assert "narrowing the scope" in description


@pytest.mark.parametrize(
    "term",
    ['deck" AND from:ceo@example.com "', 'say "hi"', "path\\", "a\\b"],
)
async def test_rejects_a_term_holding_a_quote_or_backslash(term: str) -> None:
    """Escaping these is not parsed reliably by Graph, so they are refused outright."""
    with pytest.raises(ToolError):
        await call_search(server_with(), {"term": term})
