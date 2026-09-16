from dataclasses import dataclass
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.contact_repository import ContactRepository
from outlook_mac_mcp.application.search_contacts_request import SearchContactsRequest
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.contact_repository import GraphContactRepository
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError

CONTACTS_URL = f"{GRAPH_BASE_URL}/me/contacts"


def graph_contact(contact_id: str, display_name: str, *addresses: str) -> dict[str, Any]:
    return {
        "id": contact_id,
        "displayName": display_name,
        "emailAddresses": [{"name": display_name, "address": address} for address in addresses],
    }


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def repository() -> GraphContactRepository:
    return GraphContactRepository(GraphClient(FakeTokenProvider()))


def test_satisfies_the_contact_repository_port(repository: GraphContactRepository) -> None:
    port: ContactRepository = repository

    assert port is repository


@respx.mock
def test_sends_a_name_prefix_or_email_exact_filter_with_the_count(
    repository: GraphContactRepository,
) -> None:
    route = respx.get(CONTACTS_URL).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.count": 0})
    )

    repository.search(SearchContactsRequest(term="Ana", limit=15))

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$filter"] == (
        "startswith(displayName,'Ana') or emailAddresses/any(a:a/address eq 'Ana')"
    )
    assert parameters["$count"] == "true"
    assert parameters["$top"] == "15"
    assert set(parameters["$select"].split(",")) == {"id", "displayName", "emailAddresses"}


@respx.mock
def test_maps_the_returned_contacts_and_their_addresses(repository: GraphContactRepository) -> None:
    respx.get(CONTACTS_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [graph_contact("1", "Ana Lima", "ana@x.io", "ana.lima@y.io")],
                "@odata.count": 1,
            },
        )
    )

    page = repository.search(SearchContactsRequest(term="Ana"))

    contact = page.items[0]
    assert contact.id == "1"
    assert contact.display_name == "Ana Lima"
    assert [address.address for address in contact.email_addresses] == ["ana@x.io", "ana.lima@y.io"]


@respx.mock
def test_reads_the_exact_total_from_the_count(repository: GraphContactRepository) -> None:
    respx.get(CONTACTS_URL).mock(
        return_value=httpx.Response(
            200, json={"value": [graph_contact("1", "Ana")], "@odata.count": 41}
        )
    )

    page = repository.search(SearchContactsRequest(term="Ana", limit=1))

    assert len(page.items) == 1
    assert page.total == 41
    assert page.total_is_exact is True


@respx.mock
def test_returns_empty_when_nothing_matches(repository: GraphContactRepository) -> None:
    respx.get(CONTACTS_URL).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.count": 0})
    )

    page = repository.search(SearchContactsRequest(term="zzz"))

    assert page == Page(items=(), total=0, total_is_exact=True)


@respx.mock
def test_raises_when_the_count_is_missing(repository: GraphContactRepository) -> None:
    respx.get(CONTACTS_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    with pytest.raises(GraphResponseError):
        repository.search(SearchContactsRequest(term="ana"))


@respx.mock
def test_refuses_a_term_that_could_close_the_filter_literal(
    repository: GraphContactRepository,
) -> None:
    route = respx.get(CONTACTS_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    with pytest.raises(InvalidRequestError):
        repository.search(SearchContactsRequest(term="o'neil"))

    assert route.call_count == 0
