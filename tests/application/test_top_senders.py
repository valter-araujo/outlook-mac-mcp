from datetime import UTC, datetime, timedelta

import pytest

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.top_senders import (
    DEFAULT_TOP_SENDERS,
    TopSenders,
    TopSendersRequest,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)


def make_email(email_id: str, sender: str, *, days_ago: int = 0, is_read: bool = False) -> Email:
    return Email(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address=sender),
        received_at=BASE_TIME - timedelta(days=days_ago),
        is_read=is_read,
        has_attachments=False,
        preview="",
    )


def repository_with(*emails: Email) -> InMemoryMailRepository:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(FolderName.INBOX, email)
    return repository


EVERYTHING = TopSendersRequest()


def ranked(use_case: TopSenders, request: TopSendersRequest = EVERYTHING) -> list[tuple[str, int]]:
    ranking = use_case.execute(request)
    return [(item.sender.address, item.count) for item in ranking.senders]


def test_ranks_senders_by_how_many_emails_they_sent() -> None:
    use_case = TopSenders(
        repository_with(
            make_email("1", "bo@x.io"),
            make_email("2", "ana@x.io"),
            make_email("3", "bo@x.io"),
            make_email("4", "cy@x.io"),
            make_email("5", "bo@x.io"),
            make_email("6", "ana@x.io"),
        )
    )

    assert ranked(use_case) == [("bo@x.io", 3), ("ana@x.io", 2), ("cy@x.io", 1)]


def test_breaks_ties_by_address() -> None:
    use_case = TopSenders(
        repository_with(
            make_email("1", "zed@x.io"), make_email("2", "amy@x.io"), make_email("3", "mid@x.io")
        )
    )

    assert ranked(use_case) == [("amy@x.io", 1), ("mid@x.io", 1), ("zed@x.io", 1)]


def test_reports_full_coverage_with_scanned_equal_to_total() -> None:
    use_case = TopSenders(repository_with(make_email("1", "a@x.io"), make_email("2", "b@x.io")))

    ranking = use_case.execute(TopSendersRequest())

    assert ranking.scanned == 2
    assert ranking.total == 2
    assert ranking.coverage_is_complete is True


def test_stops_at_the_ceiling_and_ranks_only_what_was_scanned() -> None:
    emails = [make_email(str(n), "old@x.io", days_ago=10 + n) for n in range(4)]
    emails += [make_email(f"new-{n}", "new@x.io", days_ago=n) for n in range(3)]
    use_case = TopSenders(repository_with(*emails), scan_ceiling=3)

    ranking = use_case.execute(TopSendersRequest())

    assert [(item.sender.address, item.count) for item in ranking.senders] == [("new@x.io", 3)]
    assert ranking.scanned == 3
    assert ranking.total == 7
    assert ranking.coverage_is_complete is False


def test_applies_the_read_state_and_date_filters() -> None:
    use_case = TopSenders(
        repository_with(
            make_email("1", "recent-read@x.io", is_read=True),
            make_email("2", "recent-unread@x.io"),
            make_email("3", "old-read@x.io", days_ago=30, is_read=True),
        )
    )
    request = TopSendersRequest(is_read=True, received_after=BASE_TIME - timedelta(days=7))

    assert ranked(use_case, request) == [("recent-read@x.io", 1)]


def test_cuts_the_ranking_at_the_limit_but_counts_everything() -> None:
    use_case = TopSenders(repository_with(*(make_email(str(n), f"{n}@x.io") for n in range(5))))

    ranking = use_case.execute(TopSendersRequest(limit=2))

    assert len(ranking.senders) == 2
    assert ranking.scanned == 5


def test_defaults_to_ten_senders_from_the_inbox() -> None:
    request = TopSendersRequest()

    assert request.limit == DEFAULT_TOP_SENDERS
    assert request.folder is FolderName.INBOX


def test_ranks_nobody_on_an_empty_folder() -> None:
    ranking = TopSenders(repository_with()).execute(TopSendersRequest())

    assert ranking.senders == ()
    assert ranking.total == 0
    assert ranking.coverage_is_complete is True


@pytest.mark.parametrize("limit", [MIN_LIMIT - 1, MAX_LIMIT + 1])
def test_rejects_a_limit_outside_the_shared_bounds(limit: int) -> None:
    with pytest.raises(InvalidRequestError):
        TopSendersRequest(limit=limit)


def test_rejects_an_inverted_date_range() -> None:
    with pytest.raises(InvalidRequestError):
        TopSendersRequest(received_after=BASE_TIME, received_before=BASE_TIME - timedelta(days=1))
