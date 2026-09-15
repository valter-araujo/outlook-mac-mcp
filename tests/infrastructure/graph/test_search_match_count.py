from dataclasses import dataclass
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.search_match_count import (
    COUNT_CEILING,
    MatchCount,
    count_search_matches,
)

INBOX_PATH = "/me/mailFolders/inbox/messages"
INBOX_URL = f"{GRAPH_BASE_URL}{INBOX_PATH}"
A_SEARCH = '"\\"deck\\""'
SECOND_PAGE_LINK = f"{INBOX_URL}?%24skip=2"
THIRD_PAGE_LINK = f"{INBOX_URL}?%24skip=4"


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def client() -> GraphClient:
    return GraphClient(FakeTokenProvider())


def ids_page(count: int, next_link: str | None = None) -> dict[str, Any]:
    page: dict[str, Any] = {"value": [{"id": str(index)} for index in range(count)]}
    if next_link is not None:
        page["@odata.nextLink"] = next_link
    return page


def first_page_route() -> respx.Route:
    """The count's own first request, told apart from a followed link by its $select."""
    return respx.get(INBOX_URL, params__contains={"$select": "id"})


def link_route(link: str, skip: str) -> respx.Route:
    return respx.get(link, params__contains={"$skip": skip})


@respx.mock
def test_counts_exactly_when_everything_fits_in_one_page(client: GraphClient) -> None:
    first_page_route().mock(return_value=httpx.Response(200, json=ids_page(3)))

    assert count_search_matches(client, INBOX_PATH, A_SEARCH) == MatchCount(3, is_exact=True)


@respx.mock
def test_counts_zero_exactly_when_nothing_matches(client: GraphClient) -> None:
    first_page_route().mock(return_value=httpx.Response(200, json=ids_page(0)))

    assert count_search_matches(client, INBOX_PATH, A_SEARCH) == MatchCount(0, is_exact=True)


@respx.mock
def test_asks_for_ids_only_and_a_page_the_size_of_the_ceiling(client: GraphClient) -> None:
    route = first_page_route().mock(return_value=httpx.Response(200, json=ids_page(0)))

    count_search_matches(client, INBOX_PATH, A_SEARCH)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$search"] == A_SEARCH
    assert parameters["$select"] == "id"
    assert parameters["$top"] == str(COUNT_CEILING)


@respx.mock
def test_follows_every_page_and_sums_them(client: GraphClient) -> None:
    first_page_route().mock(return_value=httpx.Response(200, json=ids_page(2, SECOND_PAGE_LINK)))
    link_route(SECOND_PAGE_LINK, "2").mock(
        return_value=httpx.Response(200, json=ids_page(2, THIRD_PAGE_LINK))
    )
    third = link_route(THIRD_PAGE_LINK, "4").mock(
        return_value=httpx.Response(200, json=ids_page(1))
    )

    count = count_search_matches(client, INBOX_PATH, A_SEARCH)

    assert count == MatchCount(5, is_exact=True)
    assert third.call_count == 1


@respx.mock
def test_stops_at_the_ceiling_and_reports_a_lower_bound(client: GraphClient) -> None:
    first_page_route().mock(
        return_value=httpx.Response(200, json=ids_page(COUNT_CEILING, SECOND_PAGE_LINK))
    )
    beyond = link_route(SECOND_PAGE_LINK, "2").mock(
        return_value=httpx.Response(200, json=ids_page(1))
    )

    count = count_search_matches(client, INBOX_PATH, A_SEARCH)

    assert count == MatchCount(COUNT_CEILING, is_exact=False)
    assert beyond.call_count == 0


@respx.mock
def test_reports_the_ceiling_when_the_pages_overshoot_it(client: GraphClient) -> None:
    """Graph chooses its own page sizes, so the walk can land past the ceiling."""
    first_page_route().mock(
        return_value=httpx.Response(200, json=ids_page(COUNT_CEILING - 1, SECOND_PAGE_LINK))
    )
    link_route(SECOND_PAGE_LINK, "2").mock(return_value=httpx.Response(200, json=ids_page(2)))

    count = count_search_matches(client, INBOX_PATH, A_SEARCH)

    assert count == MatchCount(COUNT_CEILING, is_exact=False)


@respx.mock
def test_a_count_landing_exactly_on_the_ceiling_is_still_exact(client: GraphClient) -> None:
    first_page_route().mock(return_value=httpx.Response(200, json=ids_page(COUNT_CEILING)))

    count = count_search_matches(client, INBOX_PATH, A_SEARCH)

    assert count == MatchCount(COUNT_CEILING, is_exact=True)


@respx.mock
def test_raises_when_a_page_has_no_value_array(client: GraphClient) -> None:
    first_page_route().mock(return_value=httpx.Response(200, json={}))

    with pytest.raises(GraphResponseError):
        count_search_matches(client, INBOX_PATH, A_SEARCH)


@respx.mock
def test_raises_when_the_next_link_is_not_a_url(client: GraphClient) -> None:
    first_page_route().mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.nextLink": 7})
    )

    with pytest.raises(GraphResponseError):
        count_search_matches(client, INBOX_PATH, A_SEARCH)
