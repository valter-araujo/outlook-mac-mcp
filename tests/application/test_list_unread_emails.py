from datetime import UTC, datetime, timedelta

import pytest

from outlook_mac_mcp.application.list_unread_emails import (
    ListUnreadEmails,
    ListUnreadEmailsRequest,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)


def make_email(email_id: str, *, is_read: bool = False, minutes_ago: int = 0) -> Email:
    return Email(
        id=email_id,
        subject="subject",
        sender=EmailAddress(address="sender@example.com"),
        received_at=BASE_TIME - timedelta(minutes=minutes_ago),
        is_read=is_read,
        has_attachments=False,
        preview="",
    )


def test_returns_empty_when_folder_has_no_unread() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("read-1", is_read=True))
    use_case = ListUnreadEmails(repository)

    result = use_case.execute(ListUnreadEmailsRequest())

    assert result.items == ()
    assert result.total == 0
    assert result.total_is_exact is True


def test_returns_only_unread_emails_from_requested_folder() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("unread-inbox"))
    repository.add(FolderName.INBOX, make_email("read-inbox", is_read=True))
    repository.add(FolderName.ARCHIVE, make_email("unread-archive"))
    use_case = ListUnreadEmails(repository)

    result = use_case.execute(ListUnreadEmailsRequest(folders=(FolderName.INBOX,)))

    assert [email.id for email in result.items] == ["unread-inbox"]


def test_reports_the_exact_number_of_unread_beyond_the_page() -> None:
    repository = InMemoryMailRepository()
    for index in range(5):
        repository.add(FolderName.INBOX, make_email(str(index), minutes_ago=index))
    repository.add(FolderName.INBOX, make_email("read", is_read=True))
    use_case = ListUnreadEmails(repository)

    result = use_case.execute(ListUnreadEmailsRequest(limit=2))

    assert len(result.items) == 2
    assert result.total == 5
    assert result.total_is_exact is True


def test_returns_newest_first_and_respects_limit() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("oldest", minutes_ago=30))
    repository.add(FolderName.INBOX, make_email("newest", minutes_ago=0))
    repository.add(FolderName.INBOX, make_email("middle", minutes_ago=10))
    use_case = ListUnreadEmails(repository)

    result = use_case.execute(ListUnreadEmailsRequest(limit=2))

    assert [email.id for email in result.items] == ["newest", "middle"]


@pytest.mark.parametrize("limit", [0, 101])
def test_rejects_limit_outside_allowed_range(limit: int) -> None:
    with pytest.raises(InvalidRequestError):
        ListUnreadEmailsRequest(limit=limit)
