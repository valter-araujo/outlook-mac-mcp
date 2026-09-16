import json
from typing import Any

import pytest
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.search_contacts_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
    SearchContactsRequest,
)
from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.infrastructure.graph.errors import NotAuthenticatedError
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.search_contacts_tool import SEARCH_CONTACTS_TOOL
from outlook_mac_mcp.interface.mcp.server import build_server
from tests.fakes.in_memory_contact_repository import InMemoryContactRepository
from tests.fakes.use_case_bundles import contacts_use_cases

pytestmark = pytest.mark.anyio


def make_contact(contact_id: str, display_name: str, *addresses: str) -> Contact:
    return Contact(
        id=contact_id,
        display_name=display_name,
        email_addresses=tuple(EmailAddress(address=address) for address in addresses),
    )


class FailingContactRepository:
    def __init__(self, error: Exception) -> None:
        self._error = error

    def search(self, request: SearchContactsRequest) -> Page[Contact]:
        raise self._error


def server_with(*contacts: Contact) -> MCPServer:
    repository = InMemoryContactRepository()
    for contact in contacts:
        repository.add(contact)
    return build_server(contacts_use_cases(repository))


async def call(server: MCPServer, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await server.call_tool(SEARCH_CONTACTS_TOOL, arguments)
    assert isinstance(result, CallToolResult)
    payload = result.structured_content
    assert isinstance(payload, dict)
    return payload


async def the_tool(server: MCPServer) -> Any:
    return next(tool for tool in await server.list_tools() if tool.name == SEARCH_CONTACTS_TOOL)


async def test_advertises_the_term_length_bounds_and_the_default_limit() -> None:
    tool = await the_tool(server_with())

    term = tool.input_schema["properties"]["term"]
    limit = tool.input_schema["properties"]["limit"]
    assert term["minLength"] == MIN_TERM_LENGTH
    assert term["maxLength"] == MAX_TERM_LENGTH
    assert limit["minimum"] == MIN_LIMIT
    assert limit["maximum"] == MAX_LIMIT


async def test_returns_a_contact_matching_a_display_name_prefix() -> None:
    server = server_with(
        make_contact("1", "Ana Lima", "ana@x.io"), make_contact("2", "Bo", "bo@x.io")
    )

    page = await call(server, {"term": "Ana"})

    assert [item["id"] for item in page["items"]] == ["1"]
    assert page["items"][0]["email_addresses"] == ["ana@x.io"]
    assert page["returned"] == 1
    assert page["total"] == 1
    assert page["total_is_exact"] is True


async def test_matches_an_exact_email_but_not_a_prefix_of_one() -> None:
    server = server_with(make_contact("1", "Ana Lima", "ana@x.io"))

    exact = await call(server, {"term": "ana@x.io"})
    prefix = await call(server, {"term": "ana@"})

    assert [item["id"] for item in exact["items"]] == ["1"]
    assert prefix["items"] == []


async def test_returns_an_empty_exact_page_when_nothing_matches() -> None:
    page = await call(server_with(), {"term": "zzz"})

    assert page == {"items": [], "returned": 0, "total": 0, "total_is_exact": True}


async def test_rejects_an_empty_term() -> None:
    with pytest.raises(ToolError):
        await call(server_with(), {"term": ""})


async def test_rejects_a_term_beyond_the_maximum_length() -> None:
    with pytest.raises(ToolError):
        await call(server_with(), {"term": "a" * (MAX_TERM_LENGTH + 1)})


async def test_rejects_a_term_holding_a_single_quote() -> None:
    with pytest.raises(ToolError):
        await call(server_with(), {"term": "o'neil"})


async def test_rejects_a_limit_outside_the_advertised_range() -> None:
    with pytest.raises(ToolError):
        await call(server_with(), {"term": "ana", "limit": MAX_LIMIT + 1})


async def test_translates_a_project_error_into_a_tool_error() -> None:
    server = build_server(
        contacts_use_cases(FailingContactRepository(NotAuthenticatedError("sign in first")))
    )

    with pytest.raises(ToolError, match="sign in first"):
        await call(server, {"term": "ana"})


async def test_says_email_matching_is_exact_and_name_matching_is_a_prefix() -> None:
    description = (await the_tool(server_with())).description or ""

    assert "PREFIX" in description
    assert "EXACT" in description
    assert '"showing N of M"' in description


async def test_logs_the_call_with_its_item_count(capsys: pytest.CaptureFixture[str]) -> None:
    configure_logging()
    server = server_with(make_contact("1", "Ana Lima", "ana@x.io"))

    await call(server, {"term": "Ana"})

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["tool_name"] == SEARCH_CONTACTS_TOOL
    assert record["item_count"] == 1
