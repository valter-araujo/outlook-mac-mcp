import json
from dataclasses import dataclass
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.email_size_scan import (
    EMAIL_SIZE_SCAN_EVENT,
    SCAN_PAGE_SIZE,
    scan_email_sizes,
)
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.large_email_mapper import MESSAGE_SIZE_PROPERTY_ID
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


def message(email_id: str, size_bytes: int) -> dict[str, Any]:
    return {
        "id": email_id,
        "subject": "subject",
        "from": {"emailAddress": {"address": "ana@example.com", "name": ""}},
        "receivedDateTime": "2026-09-14T12:00:00Z",
        "singleValueExtendedProperties": [
            {"id": MESSAGE_SIZE_PROPERTY_ID, "value": str(size_bytes)}
        ],
    }


def page(
    sized: list[tuple[str, int]], *, total: int, next_page: int | None = None
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "value": [message(email_id, size_bytes) for email_id, size_bytes in sized],
        "@odata.count": total,
    }
    if next_page is not None:
        payload["@odata.nextLink"] = f"{INBOX_URL}?%24skip={next_page}"
    return payload


def mock_first(payload: dict[str, Any]) -> respx.Route:
    return respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=payload))


def mock_page(number: int, payload: dict[str, Any]) -> respx.Route:
    return respx.get(INBOX_URL, params__contains={"$skip": str(number)}).mock(
        return_value=httpx.Response(200, json=payload)
    )


def ids_and_sizes(scan_items: Any) -> list[tuple[str, int]]:
    return [(item.id, item.size_bytes) for item in scan_items]


@respx.mock
def test_asks_for_the_size_property_expanded_with_the_filter_and_count(
    client: GraphClient,
) -> None:
    route = mock_first(page([], total=0))

    scan_email_sizes(client, INBOX_PATH, EmailFilters(is_read=False), A_CEILING)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$select"] == "id,subject,from,receivedDateTime"
    assert (
        parameters["$expand"]
        == f"singleValueExtendedProperties($filter=id eq '{MESSAGE_SIZE_PROPERTY_ID}')"
    )
    assert parameters["$top"] == str(SCAN_PAGE_SIZE)
    assert parameters["$count"] == "true"
    assert parameters["$filter"] == "isRead eq false"
    assert "$orderby" not in parameters


@respx.mock
def test_returns_every_size_of_a_single_page_with_full_coverage(client: GraphClient) -> None:
    mock_first(page([("a", 100), ("b", 9000)], total=2))

    scan = scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert ids_and_sizes(scan.items) == [("a", 100), ("b", 9000)]
    assert scan.scanned == 2
    assert scan.total == 2
    assert scan.coverage_is_complete is True


@respx.mock
def test_follows_every_page_and_keeps_the_total_from_the_first(client: GraphClient) -> None:
    mock_page(2, page([("c", 300)], total=3, next_page=3))
    last = mock_page(3, page([("d", 400)], total=3))
    mock_first(page([("a", 100), ("b", 200)], total=4, next_page=2))

    scan = scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert [item.id for item in scan.items] == ["a", "b", "c", "d"]
    assert scan.total == 4
    assert scan.coverage_is_complete is True
    assert last.call_count == 1


@respx.mock
def test_stops_mid_page_at_the_ceiling_and_reports_partial_coverage(client: GraphClient) -> None:
    beyond = mock_page(2, page([("never", 1)], total=8))
    mock_first(page([(str(n), n) for n in range(8)], total=8, next_page=2))

    scan = scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert [item.id for item in scan.items] == [str(n) for n in range(A_CEILING)]
    assert scan.total == 8
    assert scan.coverage_is_complete is False
    assert beyond.call_count == 0


@respx.mock
def test_a_last_page_that_overflows_the_ceiling_is_still_partial(client: GraphClient) -> None:
    mock_page(2, page([("d", 1), ("e", 2), ("f", 3)], total=6))
    mock_first(page([("a", 1), ("b", 2), ("c", 3)], total=6, next_page=2))

    scan = scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert scan.coverage_is_complete is False


@respx.mock
def test_a_scan_that_lands_exactly_on_the_ceiling_with_no_more_pages_is_complete(
    client: GraphClient,
) -> None:
    mock_first(page([(str(n), n) for n in range(A_CEILING)], total=A_CEILING))

    scan = scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert scan.coverage_is_complete is True


@respx.mock
def test_does_not_fetch_a_further_page_once_the_ceiling_is_reached(client: GraphClient) -> None:
    beyond = mock_page(2, page([("never", 1)], total=6))
    mock_first(page([(str(n), n) for n in range(A_CEILING)], total=6, next_page=2))

    scan = scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)

    assert scan.coverage_is_complete is False
    assert beyond.call_count == 0


@respx.mock
def test_raises_when_the_count_is_missing(client: GraphClient) -> None:
    mock_first({"value": []})

    with pytest.raises(GraphResponseError):
        scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)


@respx.mock
def test_raises_when_a_page_carries_no_size_property(client: GraphClient) -> None:
    mock_first(
        {
            "value": [
                {
                    "id": "a",
                    "subject": "s",
                    "from": None,
                    "receivedDateTime": "2026-09-14T12:00:00Z",
                    "singleValueExtendedProperties": [],
                }
            ],
            "@odata.count": 1,
        }
    )

    with pytest.raises(GraphResponseError):
        scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)


@respx.mock
def test_logs_pages_scanned_and_duration_but_never_a_subject_or_size(
    client: GraphClient, capsys: pytest.CaptureFixture[str]
) -> None:
    configure_logging()
    mock_page(2, page([("second-secret", 424242)], total=2))
    mock_first(page([("first-secret", 999999)], total=2, next_page=2))

    scan_email_sizes(client, INBOX_PATH, EmailFilters(), A_CEILING)

    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["event"] == EMAIL_SIZE_SCAN_EVENT
    assert record["pages"] == 2
    assert record["scanned"] == 2
    assert isinstance(record["duration_ms"], float)
    assert "secret" not in captured.err
    assert "999999" not in captured.err
    assert "424242" not in captured.err
