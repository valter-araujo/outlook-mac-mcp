from datetime import UTC, datetime, timedelta

from outlook_mac_mcp.application.count_emails import CountEmails
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.email_filters import EmailFilters
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


def use_case_with(*emails: Email) -> CountEmails:
    repository = InMemoryMailRepository()
    for email in emails:
        repository.add(FolderName.INBOX, email)
    return CountEmails(repository)


def test_counts_everything_in_the_folder_without_filters() -> None:
    assert (
        use_case_with(make_email("a"), make_email("b"), make_email("c")).execute(EmailFilters())
        == 3
    )


def test_counts_zero_when_the_folder_is_empty() -> None:
    assert use_case_with().execute(EmailFilters()) == 0


def test_counts_only_what_the_filters_admit() -> None:
    use_case = use_case_with(
        make_email("read-recent", is_read=True),
        make_email("unread-recent"),
        make_email("read-old", days_ago=10, is_read=True),
    )

    filters = EmailFilters(is_read=True, received_after=BASE_TIME - timedelta(days=3))
    assert use_case.execute(filters) == 1


def test_counts_the_requested_folder_only() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("inboxed"))
    repository.add(FolderName.ARCHIVE, make_email("archived"))

    assert CountEmails(repository).execute(EmailFilters(folder=FolderName.ARCHIVE)) == 1
