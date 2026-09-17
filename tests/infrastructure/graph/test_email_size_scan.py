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
from outlook_mac_mcp.infrastructure.graph.large_email_mapper import (
    MESSAGE_SIZE_PROPERTY_ID_INTEGER,
    MESSAGE_SIZE_PROPERTY_ID_LONG,
)
from outlook_mac_mcp.interface.mcp.observability import configure_logging

INBOX_PATH = "/me/mailFolders/inbox/messages"
INBOX_URL = f"{GRAPH_BASE_URL}{INBOX_PATH}"
ARCHIVE_PATH = "/me/mailFolders/archive/messages"
ARCHIVE_URL = f"{GRAPH_BASE_URL}{ARCHIVE_PATH}"
A_CEILING = 5


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def client() -> GraphClient:
    return GraphClient(FakeTokenProvider())


def message(
    email_id: str, *, size_bytes: int | None, property_id: str = MESSAGE_SIZE_PROPERTY_ID_INTEGER
) -> dict[str, Any]:
    extended_properties = (
        [] if size_bytes is None else [{"id": property_id, "value": str(size_bytes)}]
    )
    return {
        "id": email_id,
        "subject": "subject",
        "from": {"emailAddress": {"address": "ana@example.com", "name": ""}},
        "receivedDateTime": "2026-09-14T12:00:00Z",
        "singleValueExtendedProperties": extended_properties,
    }


def page(
    messages: list[dict[str, Any]], *, total: int, next_page: int | None = None
) -> dict[str, Any]:
    payload: dict[str, Any] = {"value": messages, "@odata.count": total}
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
def test_asks_for_either_property_type_expanded_with_the_filter_and_count(
    client: GraphClient,
) -> None:
    route = mock_first(page([], total=0))

    scan_email_sizes(client, [INBOX_PATH], EmailFilters(is_read=False), A_CEILING)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$select"] == "id,subject,from,receivedDateTime"
    assert parameters["$expand"] == (
        "singleValueExtendedProperties($filter=id eq "
        f"'{MESSAGE_SIZE_PROPERTY_ID_INTEGER}' or id eq '{MESSAGE_SIZE_PROPERTY_ID_LONG}')"
    )
    assert parameters["$top"] == str(SCAN_PAGE_SIZE)
    assert parameters["$count"] == "true"
    assert parameters["$filter"] == "isRead eq false"
    assert "$orderby" not in parameters


@respx.mock
def test_reads_a_message_sized_with_the_integer_typed_property(client: GraphClient) -> None:
    mock_first(
        page([message("a", size_bytes=100, property_id=MESSAGE_SIZE_PROPERTY_ID_INTEGER)], total=1)
    )

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert ids_and_sizes(scan.items) == [("a", 100)]
    assert scan.skipped == 0


@respx.mock
def test_reads_a_message_sized_with_the_long_typed_property(client: GraphClient) -> None:
    """A tenant using the Long form must be read just as reliably as one using Integer."""
    mock_first(
        page([message("a", size_bytes=200, property_id=MESSAGE_SIZE_PROPERTY_ID_LONG)], total=1)
    )

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert ids_and_sizes(scan.items) == [("a", 200)]
    assert scan.skipped == 0


@respx.mock
def test_returns_every_size_of_a_single_page_with_full_coverage(client: GraphClient) -> None:
    mock_first(page([message("a", size_bytes=100), message("b", size_bytes=9000)], total=2))

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert ids_and_sizes(scan.items) == [("a", 100), ("b", 9000)]
    assert scan.scanned == 2
    assert scan.skipped == 0
    assert scan.total == 2
    assert scan.coverage_is_complete is True


@respx.mock
def test_a_message_carrying_neither_property_is_skipped_not_raised(client: GraphClient) -> None:
    mock_first(
        page([message("sized", size_bytes=100), message("unsized", size_bytes=None)], total=2)
    )

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert ids_and_sizes(scan.items) == [("sized", 100)]
    assert scan.skipped == 1
    assert scan.scanned == 2
    assert scan.total == 2
    assert scan.coverage_is_complete is True


@respx.mock
def test_follows_every_page_and_keeps_the_total_from_the_first(client: GraphClient) -> None:
    mock_page(2, page([message("c", size_bytes=300)], total=3, next_page=3))
    last = mock_page(3, page([message("d", size_bytes=400)], total=3))
    mock_first(
        page([message("a", size_bytes=100), message("b", size_bytes=200)], total=4, next_page=2)
    )

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert [item.id for item in scan.items] == ["a", "b", "c", "d"]
    assert scan.total == 4
    assert scan.coverage_is_complete is True
    assert last.call_count == 1


@respx.mock
def test_stops_mid_page_at_the_ceiling_and_reports_partial_coverage(client: GraphClient) -> None:
    beyond = mock_page(2, page([message("never", size_bytes=1)], total=8))
    mock_first(page([message(str(n), size_bytes=n) for n in range(8)], total=8, next_page=2))

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert [item.id for item in scan.items] == [str(n) for n in range(A_CEILING)]
    assert scan.total == 8
    assert scan.coverage_is_complete is False
    assert beyond.call_count == 0


@respx.mock
def test_skipped_messages_count_toward_the_ceiling_too(client: GraphClient) -> None:
    """The ceiling bounds messages examined, sized or not, same as the sender scan."""
    unsized = [message(f"unsized-{n}", size_bytes=None) for n in range(3)]
    beyond = mock_page(2, page([message("never", size_bytes=1)], total=6))
    mock_first(
        page(
            [*unsized, message("sized", size_bytes=1), message("also-sized", size_bytes=2)],
            total=6,
            next_page=2,
        )
    )

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert scan.skipped == 3
    assert [item.id for item in scan.items] == ["sized", "also-sized"]
    assert scan.coverage_is_complete is False
    assert beyond.call_count == 0


@respx.mock
def test_a_scan_that_lands_exactly_on_the_ceiling_with_no_more_pages_is_complete(
    client: GraphClient,
) -> None:
    mock_first(page([message(str(n), size_bytes=n) for n in range(A_CEILING)], total=A_CEILING))

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert scan.scanned == A_CEILING
    assert scan.coverage_is_complete is True


@respx.mock
def test_does_not_fetch_a_further_page_once_the_ceiling_is_reached(client: GraphClient) -> None:
    beyond = mock_page(2, page([message("never", size_bytes=1)], total=6))
    mock_first(
        page([message(str(n), size_bytes=n) for n in range(A_CEILING)], total=6, next_page=2)
    )

    scan = scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert scan.coverage_is_complete is False
    assert beyond.call_count == 0


@respx.mock
def test_more_than_one_path_with_room_in_both_scans_normally(client: GraphClient) -> None:
    mock_first(page([message("a", size_bytes=10)], total=1))
    respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(200, json=page([message("b", size_bytes=20)], total=1))
    )

    scan = scan_email_sizes(client, [INBOX_PATH, ARCHIVE_PATH], EmailFilters(), A_CEILING)

    assert ids_and_sizes(scan.items) == [("a", 10), ("b", 20)]
    assert scan.total == 2
    assert scan.coverage_is_complete is True


@respx.mock
def test_more_than_one_path_shares_one_ceiling_across_both(client: GraphClient) -> None:
    """The first path alone exhausts the ceiling, so the second must never be walked --
    only its exact count, one cheap request, folds into the total.
    """
    mock_first(page([message(str(n), size_bytes=1) for n in range(A_CEILING)], total=A_CEILING))
    archive_count = respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "x"}], "@odata.count": 9})
    )

    scan = scan_email_sizes(client, [INBOX_PATH, ARCHIVE_PATH], EmailFilters(), A_CEILING)

    assert len(scan.items) == A_CEILING
    assert scan.total == A_CEILING + 9
    assert scan.coverage_is_complete is False
    parameters = dict(archive_count.calls.last.request.url.params)
    assert parameters["$top"] == "1"
    assert parameters["$select"] == "id"


@respx.mock
def test_raises_when_the_count_is_missing(client: GraphClient) -> None:
    mock_first({"value": []})

    with pytest.raises(GraphResponseError):
        scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)


@respx.mock
def test_raises_when_a_present_property_has_an_unparsable_value(client: GraphClient) -> None:
    """Presence with a garbled value is not the same as absence; it still raises."""
    mock_first(
        {
            "value": [
                {
                    "id": "a",
                    "subject": "s",
                    "from": None,
                    "receivedDateTime": "2026-09-14T12:00:00Z",
                    "singleValueExtendedProperties": [
                        {"id": MESSAGE_SIZE_PROPERTY_ID_INTEGER, "value": "not-a-number"}
                    ],
                }
            ],
            "@odata.count": 1,
        }
    )

    with pytest.raises(GraphResponseError):
        scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)


@respx.mock
def test_logs_pages_scanned_skipped_and_duration_but_never_a_subject_or_size(
    client: GraphClient, capsys: pytest.CaptureFixture[str]
) -> None:
    configure_logging()
    mock_page(2, page([message("second-secret", size_bytes=424242)], total=3, next_page=3))
    last = mock_page(3, page([message("unsized-secret", size_bytes=None)], total=3))
    mock_first(page([message("first-secret", size_bytes=999999)], total=3, next_page=2))

    scan_email_sizes(client, [INBOX_PATH], EmailFilters(), A_CEILING)

    assert last.call_count == 1
    captured = capsys.readouterr()
    assert captured.out == ""
    record = json.loads(captured.err.splitlines()[-1])
    assert record["event"] == EMAIL_SIZE_SCAN_EVENT
    assert record["pages"] == 3
    assert record["scanned"] == 3
    assert record["skipped"] == 1
    assert isinstance(record["duration_ms"], float)
    assert "secret" not in captured.err
    assert "999999" not in captured.err
    assert "424242" not in captured.err
