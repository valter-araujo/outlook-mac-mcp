"""Walking a folder's message sizes, the same shape of read as sender_scan.py.

Graph has no aggregation and no $orderby for an extended property, so ranking by size
means reading every matching message's PR_MESSAGE_SIZE, in the largest pages Graph
allows, and stopping at a ceiling so a folder of a hundred thousand messages costs a
bounded number of requests. The walk reports how far it got so the caller can say
whether the ranking covers everything, and how many examined messages carried neither
form of the property and were skipped rather than sized.
"""

from collections.abc import Mapping, Sequence
from time import perf_counter
from typing import Any

from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.email_size import EmailSize
from outlook_mac_mcp.domain.email_size_scan import EmailSizeScan
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.large_email_mapper import (
    MESSAGE_SIZE_PROPERTY_ID_INTEGER,
    MESSAGE_SIZE_PROPERTY_ID_LONG,
    to_email_size,
)
from outlook_mac_mcp.infrastructure.graph.mail_query import count_query
from outlook_mac_mcp.infrastructure.graph.pagination import read_items, read_next_link
from outlook_mac_mcp.logger import project_logger

# The largest page Graph documents for messages; fewer round trips per scan.
SCAN_PAGE_SIZE = 1000
SIZE_SCAN_SELECT = "id,subject,from,receivedDateTime"
# Either type name may be the one a given tenant actually uses for this property; a
# message only ever carries one of them, so the `or` costs nothing when both are asked.
SIZE_SCAN_EXPAND = (
    "singleValueExtendedProperties("
    f"$filter=id eq '{MESSAGE_SIZE_PROPERTY_ID_INTEGER}' or "
    f"id eq '{MESSAGE_SIZE_PROPERTY_ID_LONG}')"
)
# The smallest page Graph accepts, for a folder the ceiling has no budget left for: its
# exact total still costs one request, just not a walk through its messages.
COUNT_ONLY_PAGE_SIZE = 1
ID_ONLY_SELECT = "id"
COUNT_FIELD = "@odata.count"
EMAIL_SIZE_SCAN_EVENT = "email_size_scan"


def scan_email_sizes(
    client: GraphClient, paths: Sequence[str], filters: EmailFilters, ceiling: int
) -> EmailSizeScan:
    """Only counts and timings are logged: a subject or a size is mailbox content.

    More than one path shares one ceiling across all of them, the same bound a single
    folder's walk already respects: once messages examined (sized or skipped) reach it,
    every further folder gets only the cheap count call `total` always needs, never a
    walk through its messages. Each folder's own total is exact regardless of how much
    of it the walk actually covers -- it is one $count=true away, and that request
    happens whether or not the ceiling has room left.
    """
    started_at = perf_counter()
    sizes: list[EmailSize] = []
    skipped = 0
    total = 0
    coverage_is_complete = True
    pages = 0
    for path in paths:
        budget = ceiling - (len(sizes) + skipped)
        if budget <= 0:
            total += _count_only(client, path, filters)
            coverage_is_complete = False
            continue
        folder_sizes, folder_skipped, folder_total, folder_pages, folder_complete = (
            _walk_one_folder(client, path, filters, budget)
        )
        sizes.extend(folder_sizes)
        skipped += folder_skipped
        total += folder_total
        pages += folder_pages
        coverage_is_complete = coverage_is_complete and folder_complete
    _log(pages, len(sizes) + skipped, skipped, started_at)
    return EmailSizeScan(
        items=tuple(sizes), skipped=skipped, total=total, coverage_is_complete=coverage_is_complete
    )


def _walk_one_folder(
    client: GraphClient, path: str, filters: EmailFilters, ceiling: int
) -> tuple[list[EmailSize], int, int, int, bool]:
    """One folder's own walk, exactly as when there was only ever one folder to scan;
    `ceiling` here is however much budget this folder gets, not the tool's own ceiling.
    """
    payload = client.get(
        path,
        {
            **count_query(filters),
            "$top": SCAN_PAGE_SIZE,
            "$select": SIZE_SCAN_SELECT,
            "$expand": SIZE_SCAN_EXPAND,
        },
    )
    total = _read_count(payload)
    sizes: list[EmailSize] = []
    skipped = 0
    pages = 1
    overflowed, newly_skipped = _collect(payload, sizes, skipped, ceiling)
    skipped += newly_skipped
    next_link = read_next_link(payload)
    while next_link is not None and not overflowed and (len(sizes) + skipped) < ceiling:
        payload = client.follow(next_link)
        pages += 1
        overflowed, newly_skipped = _collect(payload, sizes, skipped, ceiling)
        skipped += newly_skipped
        next_link = read_next_link(payload)
    return sizes, skipped, total, pages, not overflowed and next_link is None


def _count_only(client: GraphClient, path: str, filters: EmailFilters) -> int:
    payload = client.get(
        path, {**count_query(filters), "$top": COUNT_ONLY_PAGE_SIZE, "$select": ID_ONLY_SELECT}
    )
    return _read_count(payload)


def _collect(
    payload: Mapping[str, Any], sizes: list[EmailSize], skipped_so_far: int, ceiling: int
) -> tuple[bool, int]:
    """Append the page's sized emails up to the ceiling, counting the rest of the page's
    unsized messages as newly skipped; report whether any item was left behind unexamined.

    The ceiling bounds messages examined, sized or not, the same way scan_senders bounds
    messages walked rather than messages that happened to carry a sender.
    """
    items = read_items(payload)
    room = ceiling - (len(sizes) + skipped_so_far)
    newly_skipped = 0
    for item in items[:room]:
        email_size = to_email_size(item)
        if email_size is None:
            newly_skipped += 1
        else:
            sizes.append(email_size)
    return len(items) > room, newly_skipped


def _read_count(payload: Mapping[str, Any]) -> int:
    count = payload.get(COUNT_FIELD)
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise GraphResponseError(f"the message collection carried no {COUNT_FIELD}")
    return count


def _log(pages: int, scanned: int, skipped: int, started_at: float) -> None:
    project_logger().info(
        EMAIL_SIZE_SCAN_EVENT,
        extra={
            "fields": {
                "pages": pages,
                "scanned": scanned,
                "skipped": skipped,
                "duration_ms": round((perf_counter() - started_at) * 1000, 3),
            }
        },
    )
