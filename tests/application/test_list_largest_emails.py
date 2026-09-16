from datetime import UTC, datetime, timedelta

import pytest

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_largest_emails import (
    DEFAULT_LARGEST_EMAILS,
    ListLargestEmails,
    ListLargestEmailsRequest,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)


def make_email(email_id: str, *, days_ago: int = 0, is_read: bool = False) -> Email:
    return Email(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address="ana@example.com"),
        received_at=BASE_TIME - timedelta(days=days_ago),
        is_read=is_read,
        has_attachments=False,
        preview="",
    )


def repository_with(*sized: tuple[Email, int | None]) -> InMemoryMailRepository:
    """A size of None stands in for a message with no size property at all."""
    repository = InMemoryMailRepository()
    for email, size_bytes in sized:
        repository.add(FolderName.INBOX, email, size_bytes=size_bytes)
    return repository


EVERYTHING = ListLargestEmailsRequest()


def ranked(
    use_case: ListLargestEmails, request: ListLargestEmailsRequest = EVERYTHING
) -> list[tuple[str, int]]:
    ranking = use_case.execute(request)
    return [(item.id, item.size_bytes) for item in ranking.items]


def test_ranks_emails_largest_first() -> None:
    use_case = ListLargestEmails(
        repository_with(
            (make_email("small"), 100),
            (make_email("huge", days_ago=1), 90_000),
            (make_email("mid", days_ago=2), 5_000),
        )
    )

    assert ranked(use_case) == [("huge", 90_000), ("mid", 5_000), ("small", 100)]


def test_keeps_scan_order_among_equal_sizes() -> None:
    use_case = ListLargestEmails(
        repository_with(
            (make_email("newest"), 100),
            (make_email("older", days_ago=1), 100),
        )
    )

    assert [item_id for item_id, _ in ranked(use_case)] == ["newest", "older"]


def test_reports_full_coverage_with_scanned_equal_to_total() -> None:
    use_case = ListLargestEmails(repository_with((make_email("1"), 10), (make_email("2"), 20)))

    ranking = use_case.execute(ListLargestEmailsRequest())

    assert ranking.scanned == 2
    assert ranking.skipped == 0
    assert ranking.total == 2
    assert ranking.coverage_is_complete is True


def test_an_email_with_no_known_size_is_skipped_not_ranked() -> None:
    use_case = ListLargestEmails(
        repository_with((make_email("sized"), 100), (make_email("unsized", days_ago=1), None))
    )

    ranking = use_case.execute(ListLargestEmailsRequest())

    assert [item.id for item in ranking.items] == ["sized"]
    assert ranking.skipped == 1
    assert ranking.scanned == 2
    assert ranking.total == 2


def test_stops_at_the_ceiling_and_ranks_only_what_was_scanned() -> None:
    small = [(make_email(str(n), days_ago=10 + n), 10) for n in range(4)]
    large_recent = [(make_email(f"new-{n}", days_ago=n), 9_000) for n in range(3)]
    use_case = ListLargestEmails(repository_with(*small, *large_recent), scan_ceiling=3)

    ranking = use_case.execute(ListLargestEmailsRequest())

    assert [item.id for item in ranking.items] == ["new-0", "new-1", "new-2"]
    assert ranking.scanned == 3
    assert ranking.total == 7
    assert ranking.coverage_is_complete is False


def test_applies_the_read_state_and_date_filters() -> None:
    use_case = ListLargestEmails(
        repository_with(
            (make_email("recent-read", is_read=True), 500),
            (make_email("recent-unread"), 900),
            (make_email("old-read", days_ago=30, is_read=True), 700),
        )
    )
    request = ListLargestEmailsRequest(is_read=True, received_after=BASE_TIME - timedelta(days=7))

    assert ranked(use_case, request) == [("recent-read", 500)]


def test_cuts_the_ranking_at_the_limit_but_counts_everything() -> None:
    use_case = ListLargestEmails(repository_with(*((make_email(str(n)), n) for n in range(5))))

    ranking = use_case.execute(ListLargestEmailsRequest(limit=2))

    assert len(ranking.items) == 2
    assert ranking.scanned == 5


def test_defaults_to_ten_largest_from_the_inbox() -> None:
    request = ListLargestEmailsRequest()

    assert request.limit == DEFAULT_LARGEST_EMAILS
    assert request.folder is FolderName.INBOX


def test_ranks_nothing_on_an_empty_folder() -> None:
    ranking = ListLargestEmails(repository_with()).execute(ListLargestEmailsRequest())

    assert ranking.items == ()
    assert ranking.total == 0
    assert ranking.coverage_is_complete is True


@pytest.mark.parametrize("limit", [MIN_LIMIT - 1, MAX_LIMIT + 1])
def test_rejects_a_limit_outside_the_shared_bounds(limit: int) -> None:
    with pytest.raises(InvalidRequestError):
        ListLargestEmailsRequest(limit=limit)
