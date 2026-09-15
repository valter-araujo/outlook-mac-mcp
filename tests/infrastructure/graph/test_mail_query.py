from datetime import UTC, datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.sort_order import SortOrder
from outlook_mac_mcp.infrastructure.graph.errors import MalformedQueryError
from outlook_mac_mcp.infrastructure.graph.mail_query import (
    EARLIEST_RECEIVED_AT,
    count_query,
    ensure_sort_property_leads,
    filter_clauses,
    list_query,
)

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
LOCAL_MIDNIGHT = datetime(2026, 9, 1, tzinfo=SAO_PAULO)
SENTINEL = f"receivedDateTime ge {EARLIEST_RECEIVED_AT}"


def test_lists_everything_newest_first_without_a_filter() -> None:
    query = list_query(EmailFilters(), SortOrder.NEWEST)

    assert query == {"$orderby": "receivedDateTime desc", "$count": "true"}


def test_lists_oldest_first_when_asked() -> None:
    assert list_query(EmailFilters(), SortOrder.OLDEST)["$orderby"] == "receivedDateTime asc"


def test_renders_received_after_as_an_inclusive_utc_bound() -> None:
    clauses = filter_clauses(EmailFilters(received_after=LOCAL_MIDNIGHT))

    assert clauses == ["receivedDateTime ge 2026-09-01T03:00:00Z"]


def test_renders_received_before_as_an_exclusive_utc_bound() -> None:
    clauses = filter_clauses(EmailFilters(received_before=LOCAL_MIDNIGHT))

    assert clauses == ["receivedDateTime lt 2026-09-01T03:00:00Z"]


def test_keeps_fractional_seconds_in_a_bound() -> None:
    bound = datetime(2026, 9, 1, 12, 0, 0, 250000, tzinfo=UTC)

    assert filter_clauses(EmailFilters(received_after=bound)) == [
        "receivedDateTime ge 2026-09-01T12:00:00.250000Z"
    ]


@pytest.mark.parametrize(("is_read", "literal"), [(True, "true"), (False, "false")])
def test_renders_the_read_state(is_read: bool, literal: str) -> None:
    assert filter_clauses(EmailFilters(is_read=is_read)) == [f"isRead eq {literal}"]


@pytest.mark.parametrize(("has_attachments", "literal"), [(True, "true"), (False, "false")])
def test_renders_the_attachment_state(has_attachments: bool, literal: str) -> None:
    clauses = filter_clauses(EmailFilters(has_attachments=has_attachments))

    assert clauses == [f"hasAttachments eq {literal}"]


def test_renders_the_sender_as_an_exact_quoted_address() -> None:
    clauses = filter_clauses(EmailFilters(sender="ana@example.com"))

    assert clauses == ["from/emailAddress/address eq 'ana@example.com'"]


def test_combines_every_filter_with_the_dates_first() -> None:
    filters = EmailFilters(
        sender="ana@example.com",
        has_attachments=True,
        is_read=False,
        received_before=LOCAL_MIDNIGHT + timedelta(days=1),
        received_after=LOCAL_MIDNIGHT,
    )

    assert list_query(filters, SortOrder.NEWEST)["$filter"] == (
        "receivedDateTime ge 2026-09-01T03:00:00Z and "
        "receivedDateTime lt 2026-09-02T03:00:00Z and "
        "isRead eq false and "
        "hasAttachments eq true and "
        "from/emailAddress/address eq 'ana@example.com'"
    )


def test_leads_with_an_epoch_bound_when_filtering_without_a_date() -> None:
    query = list_query(EmailFilters(is_read=False, sender="ana@example.com"), SortOrder.NEWEST)

    assert query["$filter"] == (
        f"{SENTINEL} and isRead eq false and from/emailAddress/address eq 'ana@example.com'"
    )


def test_adds_no_sentinel_when_a_date_bound_already_leads() -> None:
    query = list_query(EmailFilters(is_read=False, received_after=LOCAL_MIDNIGHT), SortOrder.NEWEST)

    assert SENTINEL not in query["$filter"]
    assert query["$filter"].startswith("receivedDateTime ge 2026-09-01T03:00:00Z")


def test_always_asks_for_the_count() -> None:
    assert list_query(EmailFilters(is_read=True), SortOrder.NEWEST)["$count"] == "true"
    assert count_query(EmailFilters(is_read=True))["$count"] == "true"


def test_count_sends_the_same_filter_without_an_order_or_a_sentinel() -> None:
    query = count_query(EmailFilters(is_read=False, sender="ana@example.com"))

    assert query["$filter"] == "isRead eq false and from/emailAddress/address eq 'ana@example.com'"
    assert "$orderby" not in query


def test_count_without_filters_sends_no_filter() -> None:
    assert count_query(EmailFilters()) == {"$count": "true"}


def test_the_ordering_rule_accepts_an_empty_filter_and_a_leading_date() -> None:
    ensure_sort_property_leads([])
    ensure_sort_property_leads(["receivedDateTime ge 2026-09-01T00:00:00Z", "isRead eq true"])


def test_the_ordering_rule_rejects_a_filter_that_does_not_lead_with_the_sort_property() -> None:
    with pytest.raises(MalformedQueryError, match="receivedDateTime"):
        ensure_sort_property_leads(["isRead eq true", "receivedDateTime ge 2026-09-01T00:00:00Z"])
