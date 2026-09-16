from datetime import UTC, datetime

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.email_size import EmailSize
from outlook_mac_mcp.domain.email_size_ranking import rank_by_size

RECEIVED_AT = datetime(2026, 9, 14, tzinfo=UTC)


def make(email_id: str, size_bytes: int) -> EmailSize:
    return EmailSize(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address="ana@example.com"),
        received_at=RECEIVED_AT,
        size_bytes=size_bytes,
    )


def test_ranks_largest_first() -> None:
    items = [make("small", 100), make("large", 9000), make("mid", 500)]

    ranked = rank_by_size(items, limit=10)

    assert [item.id for item in ranked] == ["large", "mid", "small"]


def test_keeps_scan_order_among_equal_sizes() -> None:
    items = [make("first", 100), make("second", 100), make("third", 100)]

    ranked = rank_by_size(items, limit=10)

    assert [item.id for item in ranked] == ["first", "second", "third"]


def test_cuts_the_ranking_at_the_limit() -> None:
    items = [make(str(n), n) for n in range(5)]

    assert len(rank_by_size(items, limit=2)) == 2
    assert [item.id for item in rank_by_size(items, limit=2)] == ["4", "3"]


def test_ranks_nothing_when_nothing_was_scanned() -> None:
    assert rank_by_size([], limit=10) == ()
