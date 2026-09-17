"""Walking a folder's senders, which is the one Graph read that touches many messages.

Graph has no aggregation, so ranking senders means reading who sent every message that
matches, `from` only, in the largest pages it allows, and stopping at a ceiling so a
folder of a hundred thousand messages costs a bounded number of requests. The walk
reports how far it got so the caller can say whether the ranking covers everything.
"""

from collections.abc import Mapping, Sequence
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
# The smallest page Graph accepts, for a folder the ceiling has no budget left for: its
# exact total still costs one request, just not a walk through its messages.
COUNT_ONLY_PAGE_SIZE = 1
ID_ONLY_SELECT = "id"
COUNT_FIELD = "@odata.count"
SENDER_SCAN_EVENT = "sender_scan"


def scan_senders(
    client: GraphClient, paths: Sequence[str], filters: EmailFilters, ceiling: int
) -> SenderScan:
    """Only counts and timings are logged: an address is mailbox content.

    More than one path shares one ceiling across all of them, the same bound a single
    folder's walk already respects, rather than a separate ceiling per folder: once the
    senders collected so far reach it, every further folder gets only the cheap count
    call `total` always needs, never a walk through its messages. Each folder's own
    total is exact regardless of how much of it the walk actually covers -- it is one
    $count=true away, and that request happens whether or not the ceiling has room left.
    """
    started_at = perf_counter()
    senders: list[EmailAddress] = []
    total = 0
    coverage_is_complete = True
    pages = 0
    for path in paths:
        budget = ceiling - len(senders)
        if budget <= 0:
            total += _count_only(client, path, filters)
            coverage_is_complete = False
            continue
        folder_senders, folder_total, folder_pages, folder_complete = _walk_one_folder(
            client, path, filters, budget
        )
        senders.extend(folder_senders)
        total += folder_total
        pages += folder_pages
        coverage_is_complete = coverage_is_complete and folder_complete
    _log(pages, len(senders), started_at)
    return SenderScan(
        senders=tuple(senders), total=total, coverage_is_complete=coverage_is_complete
    )


def _walk_one_folder(
    client: GraphClient, path: str, filters: EmailFilters, ceiling: int
) -> tuple[list[EmailAddress], int, int, bool]:
    """One folder's own walk, exactly as when there was only ever one folder to scan;
    `ceiling` here is however much budget this folder gets, not the tool's own ceiling.
    """
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
    return senders, total, pages, not overflowed and next_link is None


def _count_only(client: GraphClient, path: str, filters: EmailFilters) -> int:
    payload = client.get(
        path, {**count_query(filters), "$top": COUNT_ONLY_PAGE_SIZE, "$select": ID_ONLY_SELECT}
    )
    return _read_count(payload)


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
