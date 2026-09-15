from datetime import UTC, datetime, timedelta

import pytest

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.application.search_emails_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
    SearchEmailsRequest,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

BASE_TIME = datetime(2026, 9, 14, 12, 0, tzinfo=UTC)


def make_email(email_id: str, *, subject: str = "", minutes_ago: int = 0) -> Email:
    return Email(
        id=email_id,
        subject=subject,
        sender=EmailAddress(address="ana@example.com"),
        received_at=BASE_TIME - timedelta(minutes=minutes_ago),
        is_read=False,
        has_attachments=False,
        preview="",
    )


def test_returns_the_emails_matching_the_term() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("hit", subject="Quarterly review"))
    repository.add(FolderName.INBOX, make_email("miss", subject="Lunch"))
    use_case = SearchEmails(repository)

    result = use_case.execute(SearchEmailsRequest(term="quarterly"))

    assert [email.id for email in result.items] == ["hit"]


def test_searches_only_the_requested_folder() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("inboxed", subject="deck"))
    repository.add(FolderName.ARCHIVE, make_email("archived", subject="deck"))
    use_case = SearchEmails(repository)

    result = use_case.execute(SearchEmailsRequest(term="deck", folder=FolderName.ARCHIVE))

    assert [email.id for email in result.items] == ["archived"]


def test_returns_empty_when_nothing_matches() -> None:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, make_email("one", subject="Lunch"))

    result = SearchEmails(repository).execute(SearchEmailsRequest(term="payroll"))

    assert result.items == ()
    assert result.total == 0
    assert result.total_is_exact is True


def test_respects_the_limit() -> None:
    repository = InMemoryMailRepository()
    for index in range(5):
        repository.add(FolderName.INBOX, make_email(str(index), subject="deck"))

    result = SearchEmails(repository).execute(SearchEmailsRequest(term="deck", limit=2))

    assert len(result.items) == 2


def test_reports_the_exact_number_of_matches_beyond_the_page() -> None:
    repository = InMemoryMailRepository()
    for index in range(5):
        repository.add(FolderName.INBOX, make_email(str(index), subject="deck"))
    repository.add(FolderName.INBOX, make_email("other", subject="lunch"))

    result = SearchEmails(repository).execute(SearchEmailsRequest(term="deck", limit=2))

    assert result.total == 5
    assert result.total_is_exact is True


def test_defaults_to_the_inbox_and_the_shared_limit() -> None:
    request = SearchEmailsRequest(term="deck")

    assert request.folder is FolderName.INBOX
    assert request.limit == DEFAULT_LIMIT


@pytest.mark.parametrize("length", [MIN_TERM_LENGTH, MAX_TERM_LENGTH])
def test_accepts_a_term_at_the_edge_of_the_allowed_length(length: int) -> None:
    assert SearchEmailsRequest(term="a" * length).term


@pytest.mark.parametrize("length", [MIN_TERM_LENGTH - 1, MAX_TERM_LENGTH + 1])
def test_rejects_a_term_outside_the_allowed_length(length: int) -> None:
    with pytest.raises(InvalidRequestError):
        SearchEmailsRequest(term="a" * length)


@pytest.mark.parametrize("limit", [MIN_LIMIT - 1, MAX_LIMIT + 1])
def test_rejects_a_limit_outside_the_shared_bounds(limit: int) -> None:
    with pytest.raises(InvalidRequestError):
        SearchEmailsRequest(term="deck", limit=limit)
