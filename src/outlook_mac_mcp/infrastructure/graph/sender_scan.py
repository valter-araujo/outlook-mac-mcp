"""Walking a folder's senders, which is the one Graph read that touches many messages.

Graph has no aggregation, so ranking senders means reading who sent every message that
matches, `from` only, in the largest pages it allows, and stopping at a ceiling so a
folder of a hundred thousand messages costs a bounded number of requests. The walk
reports how far it got so the caller can say whether the ranking covers everything.
"""

from collections.abc import Mapping
from time import perf_counter
from typing import Any

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.sender_scan import SenderScan
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import to_email_address
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.mail_query import count_query
from outlook_mac_mcp.infrastructure.graph.pagination import read_items, read_next_link
from outlook_mac_mcp.logger import project_logger

# The largest page Graph documents for messages; fewer round trips per scan.
SCAN_PAGE_SIZE = 1000
SENDER_ONLY_SELECT = "from"
COUNT_FIELD = "@odata.count"
SENDER_SCAN_EVENT = "sender_scan"


def scan_senders(client: GraphClient, path: str, filters: EmailFilters, ceiling: int) -> SenderScan:
    """Only counts and timings are logged: an address is mailbox content."""
    started_at = perf_counter()
    payload = client.get(
        path, {**count_query(filters), "$top": SCAN_PAGE_SIZE, "$select": SENDER_ONLY_SELECT}
    )
    total = _read_count(payload)
    senders: list[EmailAddress] = []
    pages = 1
    overflowed = _collect(payload, senders, ceiling)
    next_link = read_next_link(payload)
    while next_link is not None and not overflowed and len(senders) < ceiling:
        payload = client.follow(next_link)
        pages += 1
        overflowed = _collect(payload, senders, ceiling)
        next_link = read_next_link(payload)
    _log(pages, len(senders), started_at)
    return SenderScan(
        senders=tuple(senders),
        total=total,
        coverage_is_complete=not overflowed and next_link is None,
    )


def _collect(payload: Mapping[str, Any], senders: list[EmailAddress], ceiling: int) -> bool:
    """Append the page's senders up to the ceiling; report whether any were left behind."""
    items = read_items(payload)
    room = ceiling - len(senders)
    senders.extend(to_email_address(item.get("from")) for item in items[:room])
    return len(items) > room


def _read_count(payload: Mapping[str, Any]) -> int:
    count = payload.get(COUNT_FIELD)
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise GraphResponseError(f"the message collection carried no {COUNT_FIELD}")
    return count


def _log(pages: int, scanned: int, started_at: float) -> None:
    project_logger().info(
        SENDER_SCAN_EVENT,
        extra={
            "fields": {
                "pages": pages,
                "scanned": scanned,
                "duration_ms": round((perf_counter() - started_at) * 1000, 3),
            }
        },
    )
