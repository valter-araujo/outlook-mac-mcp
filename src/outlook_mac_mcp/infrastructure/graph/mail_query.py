"""Building the one $filter behind list_emails and count_emails.

Graph documents a rule for messages: when $filter and $orderby are combined, every
property in $orderby must also appear in $filter, first and in the same order. The
listing sorts by receivedDateTime, so its $filter always leads with receivedDateTime.
When the caller gave no date bound but did give other filters, a lower bound at the
epoch is prepended: it excludes nothing and keeps the rule satisfied.

Nothing user-supplied is spliced in as text. Booleans come from a two-entry table, dates
are rendered from aware datetimes, and the sender is an address that EmailFilters has
already refused to accept with a quote in it, so no value can close the OData literal.
"""

from collections.abc import Sequence
from datetime import UTC, datetime

from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.sort_order import SortOrder
from outlook_mac_mcp.infrastructure.graph.errors import MalformedQueryError

SORT_PROPERTY = "receivedDateTime"
SENDER_PROPERTY = "from/emailAddress/address"
EARLIEST_RECEIVED_AT = "1970-01-01T00:00:00Z"
DIRECTION_BY_SORT = {SortOrder.NEWEST: "desc", SortOrder.OLDEST: "asc"}
BOOLEAN_LITERAL = {True: "true", False: "false"}


def list_query(filters: EmailFilters, sort: SortOrder) -> dict[str, str]:
    clauses = filter_clauses(filters)
    if clauses and not _leads_with_sort_property(clauses):
        clauses.insert(0, f"{SORT_PROPERTY} ge {EARLIEST_RECEIVED_AT}")
    ensure_sort_property_leads(clauses)
    query = {"$orderby": f"{SORT_PROPERTY} {DIRECTION_BY_SORT[sort]}", "$count": "true"}
    return _with_filter(query, clauses)


def count_query(filters: EmailFilters) -> dict[str, str]:
    """No $orderby, so the ordering rule does not apply and no sentinel is needed."""
    return _with_filter({"$count": "true"}, filter_clauses(filters))


def filter_clauses(filters: EmailFilters) -> list[str]:
    """The clauses in the order the listing needs them: the sort property's first."""
    clauses: list[str] = []
    if filters.received_after is not None:
        clauses.append(f"{SORT_PROPERTY} ge {_utc(filters.received_after)}")
    if filters.received_before is not None:
        clauses.append(f"{SORT_PROPERTY} lt {_utc(filters.received_before)}")
    if filters.is_read is not None:
        clauses.append(f"isRead eq {BOOLEAN_LITERAL[filters.is_read]}")
    if filters.has_attachments is not None:
        clauses.append(f"hasAttachments eq {BOOLEAN_LITERAL[filters.has_attachments]}")
    if filters.sender is not None:
        clauses.append(f"{SENDER_PROPERTY} eq '{filters.sender}'")
    return clauses


def ensure_sort_property_leads(clauses: Sequence[str]) -> None:
    """Refuse a $filter that Graph is documented to reject alongside the $orderby."""
    if clauses and not _leads_with_sort_property(clauses):
        raise MalformedQueryError(
            f"a filter combined with an order by {SORT_PROPERTY} must lead with {SORT_PROPERTY}"
        )


def _leads_with_sort_property(clauses: Sequence[str]) -> bool:
    return clauses[0].startswith(f"{SORT_PROPERTY} ")


def _with_filter(query: dict[str, str], clauses: Sequence[str]) -> dict[str, str]:
    if clauses:
        query["$filter"] = " and ".join(clauses)
    return query


def _utc(instant: datetime) -> str:
    """ISO 8601 in UTC with a Z suffix, which is the form Graph documents for dates."""
    return instant.astimezone(UTC).isoformat().replace("+00:00", "Z")
