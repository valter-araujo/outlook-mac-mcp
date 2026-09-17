from datetime import UTC, datetime, timedelta

import pytest

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_emails import ListEmails
from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.sort_order import SortOrder
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)


def make_email(
    email_id: str,
    *,
    days_ago: int = 0,
    is_read: bool = False,
    sender: str = "ana@example.com",
    has_attachments: bool = False,
) -> Email:
    return Email(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address=sender),
        received_at=BASE_TIME - timedelta(days=days_ago),
        is_read=is_read,
        has_attachments=has_attachments,
        preview="",
    )


def use_case_with(*emails: Email, folder: FolderName = FolderName.INBOX) -> ListEmails:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(folder, email)
    return ListEmails(repository)


def ids(use_case: ListEmails, request: ListEmailsRequest) -> list[str]:
    return [email.id for email in use_case.execute(request).items]


def test_lists_everything_newest_first_by_default() -> None:
    use_case = use_case_with(
        make_email("old", days_ago=2), make_email("new"), make_email("mid", days_ago=1)
    )

    assert ids(use_case, ListEmailsRequest()) == ["new", "mid", "old"]


def test_lists_oldest_first_when_asked() -> None:
    use_case = use_case_with(make_email("old", days_ago=2), make_email("new"))

    assert ids(use_case, ListEmailsRequest(sort=SortOrder.OLDEST)) == ["old", "new"]


def test_reports_the_exact_total_beyond_the_page() -> None:
    use_case = use_case_with(*(make_email(str(index), days_ago=index) for index in range(5)))

    page = use_case.execute(ListEmailsRequest(limit=2))

    assert len(page.items) == 2
    assert page.total == 5
    assert page.total_is_exact is True


def test_filters_by_read_state() -> None:
    use_case = use_case_with(make_email("unread"), make_email("read", is_read=True))

    assert ids(use_case, ListEmailsRequest(filters=EmailFilters(is_read=True))) == ["read"]
    assert ids(use_case, ListEmailsRequest(filters=EmailFilters(is_read=False))) == ["unread"]


def test_filters_by_sender_exactly_ignoring_case() -> None:
    use_case = use_case_with(
        make_email("from-ana", sender="Ana@Example.com"),
        make_email("from-bo", sender="bo@example.com"),
        make_email("from-ana-lima", sender="ana.lima@example.com"),
    )

    assert ids(use_case, ListEmailsRequest(filters=EmailFilters(sender="ana@example.com"))) == [
        "from-ana"
    ]


def test_filters_by_attachments() -> None:
    use_case = use_case_with(make_email("plain"), make_email("attached", has_attachments=True))

    filters = EmailFilters(has_attachments=True)
    assert ids(use_case, ListEmailsRequest(filters=filters)) == ["attached"]


def test_received_after_is_inclusive_and_received_before_exclusive() -> None:
    use_case = use_case_with(
        make_email("three-days", days_ago=3),
        make_email("two-days", days_ago=2),
        make_email("one-day", days_ago=1),
    )
    filters = EmailFilters(
        received_after=BASE_TIME - timedelta(days=2), received_before=BASE_TIME - timedelta(days=1)
    )

    assert ids(use_case, ListEmailsRequest(filters=filters)) == ["two-days"]


def test_combines_every_filter() -> None:
    use_case = use_case_with(
        make_email("hit", days_ago=1, is_read=True, sender="bo@example.com", has_attachments=True),
        make_email("wrong-sender", days_ago=1, is_read=True, has_attachments=True),
        make_email(
            "too-old", days_ago=5, is_read=True, sender="bo@example.com", has_attachments=True
        ),
        make_email("unread", days_ago=1, sender="bo@example.com", has_attachments=True),
    )
    filters = EmailFilters(
        is_read=True,
        sender="bo@example.com",
        received_after=BASE_TIME - timedelta(days=2),
        has_attachments=True,
    )

    assert ids(use_case, ListEmailsRequest(filters=filters)) == ["hit"]


def test_reads_the_requested_folder_only() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("inboxed"))
    repository.add(FolderName.ARCHIVE, make_email("archived"))

    page = ListEmails(repository).execute(
        ListEmailsRequest(filters=EmailFilters(folders=(FolderName.ARCHIVE,)))
    )

    assert [email.id for email in page.items] == ["archived"]


def test_returns_an_empty_exact_page_when_nothing_matches() -> None:
    page = use_case_with().execute(ListEmailsRequest())

    assert page.items == ()
    assert page.total == 0
    assert page.total_is_exact is True


@pytest.mark.parametrize("limit", [MIN_LIMIT - 1, MAX_LIMIT + 1])
def test_rejects_a_limit_outside_the_shared_bounds(limit: int) -> None:
    with pytest.raises(InvalidRequestError):
        ListEmailsRequest(limit=limit)
