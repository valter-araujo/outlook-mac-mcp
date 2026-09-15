import json
from dataclasses import dataclass
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.sender_scan import (
    SCAN_PAGE_SIZE,
    SENDER_SCAN_EVENT,
    scan_senders,
)
from outlook_mac_mcp.interface.mcp.observability import configure_logging

INBOX_PATH = "/me/mailFolders/inbox/messages"
INBOX_URL = f"{GRAPH_BASE_URL}{INBOX_PATH}"
A_CEILING = 5


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def client() -> GraphClient:
    return GraphClient(FakeTokenProvider())


def message_from(address: str) -> dict[str, Any]:
    return {"from": {"emailAddress": {"address": address, "name": ""}}}


def page(addresses: list[str], *, total: int, next_page: int | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {"value": [message_from(a) for a in addresses], "@odata.count": total}
    if next_page is not None:
        payload["@odata.nextLink"] = f"{INBOX_URL}?%24skip={next_page}"
    return payload


def mock_first(payload: dict[str, Any]) -> respx.Route:
    return respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=payload))


def mock_page(number: int, payload: dict[str, Any]) -> respx.Route:
    return respx.get(INBOX_URL, params__contains={"$skip": str(number)}).mock(
        return_value=httpx.Response(200, json=payload)
    )


def addresses_of(senders: tuple[EmailAddress, ...]) -> list[str]:
    return [sender.address for sender in senders]


@respx.mock
def test_asks_for_senders_only_in_the_largest_page_with_the_filter_and_count(
    client: GraphClient,
) -> None:
    route = mock_first(page([], total=0))

    scan_senders(client, INBOX_PATH, EmailFilters(is_read=False), A_CEILING)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$select"] == "from"
    assert parameters["$top"] == str(SCAN_PAGE_SIZE)
    assert parameters["$count"] == "true"
    assert parameters["$filter"] == "isRead eq false"
    assert "$orderby" not in parameters


@respx.mock
def test_returns_every_sender_of_a_single_page_with_full_coverage(client: GraphClient) -> None:
    mock_first(page(["a@x.io", "b@x.io", "a@x.io"], total=3))

    scan = scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert addresses_of(scan.senders) == ["a@x.io", "b@x.io", "a@x.io"]
    assert scan.scanned == 3
    assert scan.total == 3
    assert scan.coverage_is_complete is True


@respx.mock
def test_follows_every_page_and_keeps_the_total_from_the_first(client: GraphClient) -> None:
    mock_page(2, page(["c@x.io"], total=3, next_page=3))
    last = mock_page(3, page(["d@x.io"], total=3))
    mock_first(page(["a@x.io", "b@x.io"], total=4, next_page=2))

    scan = scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert addresses_of(scan.senders) == ["a@x.io", "b@x.io", "c@x.io", "d@x.io"]
    assert scan.total == 4
    assert scan.coverage_is_complete is True
    assert last.call_count == 1


@respx.mock
def test_stops_mid_page_at_the_ceiling_and_reports_partial_coverage(client: GraphClient) -> None:
    beyond = mock_page(2, page(["never@x.io"], total=8))
    mock_first(page([f"{n}@x.io" for n in range(8)], total=8, next_page=2))

    scan = scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert addresses_of(scan.senders) == [f"{n}@x.io" for n in range(A_CEILING)]
    assert scan.total == 8
    assert scan.coverage_is_complete is False
    assert beyond.call_count == 0


@respx.mock
def test_a_last_page_that_overflows_the_ceiling_is_still_partial(client: GraphClient) -> None:
    mock_page(2, page(["d@x.io", "e@x.io", "f@x.io"], total=6))
    mock_first(page(["a@x.io", "b@x.io", "c@x.io"], total=6, next_page=2))

    scan = scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert scan.coverage_is_complete is False


@respx.mock
def test_a_scan_that_lands_exactly_on_the_ceiling_with_no_more_pages_is_complete(
    client: GraphClient,
) -> None:
    mock_first(page([f"{n}@x.io" for n in range(A_CEILING)], total=A_CEILING))

    scan = scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert scan.coverage_is_complete is True


@respx.mock
def test_does_not_fetch_a_further_page_once_the_ceiling_is_reached(client: GraphClient) -> None:
    beyond = mock_page(2, page(["never@x.io"], total=6))
    mock_first(page([f"{n}@x.io" for n in range(A_CEILING)], total=6, next_page=2))

    scan = scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.coverage_is_complete is False
    assert beyond.call_count == 0


@respx.mock
def test_scans_a_message_with_no_sender_as_an_empty_address(client: GraphClient) -> None:
    mock_first({"value": [{"id": "draft"}, message_from("a@x.io")], "@odata.count": 2})

    scan = scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert addresses_of(scan.senders) == ["", "a@x.io"]


@respx.mock
def test_raises_when_the_count_is_missing(client: GraphClient) -> None:
    mock_first({"value": []})

    with pytest.raises(GraphResponseError):
        scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)


@respx.mock
def test_logs_pages_scanned_and_duration_but_never_an_address(
    client: GraphClient, capsys: pytest.CaptureFixture[str]
) -> None:
    configure_logging()
    mock_page(2, page(["secret@x.io"], total=2))
    mock_first(page(["ceo@x.io"], total=2, next_page=2))

    scan_senders(client, INBOX_PATH, EmailFilters(), A_CEILING)

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["event"] == SENDER_SCAN_EVENT
    assert record["pages"] == 2
    assert record["scanned"] == 2
    assert isinstance(record["duration_ms"], float)
    assert "x.io" not in captured.err
