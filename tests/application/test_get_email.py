from datetime import UTC, datetime

import pytest

from outlook_mac_mcp.application.get_email import (
    MAX_EMAIL_ID_LENGTH,
    GetEmail,
    GetEmailRequest,
)
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import EmailNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from tests.fakes.in_memory_mail_repository import InMemoryMailRepository

AN_EMAIL = Email(
    id="AAMkAGI2",
    subject="Quarterly review",
    sender=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
    received_at=datetime(2026, 9, 14, 12, 30, tzinfo=UTC),
    is_read=False,
    has_attachments=False,
    preview="Attached is the deck",
)


def use_case_with(email: Email, body: str) -> GetEmail:
    repository = InMemoryMailRepository()
    repository.add(FolderName.INBOX, email, body)
    return GetEmail(repository)


def test_returns_the_email_with_its_body() -> None:
    use_case = use_case_with(AN_EMAIL, "The full text of the message.")

    detail = use_case.execute(GetEmailRequest(email_id="AAMkAGI2"))

    assert detail.email == AN_EMAIL
    assert detail.body == "The full text of the message."


def test_returns_an_empty_body_when_the_message_has_none() -> None:
    use_case = use_case_with(AN_EMAIL, "")

    assert use_case.execute(GetEmailRequest(email_id="AAMkAGI2")).body == ""


def test_raises_when_the_mailbox_holds_no_such_email() -> None:
    use_case = use_case_with(AN_EMAIL, "body")

    with pytest.raises(EmailNotFoundError):
        use_case.execute(GetEmailRequest(email_id="missing"))


def test_rejects_an_empty_id() -> None:
    with pytest.raises(InvalidRequestError):
        GetEmailRequest(email_id="")


def test_accepts_an_id_at_the_maximum_length() -> None:
    assert GetEmailRequest(email_id="a" * MAX_EMAIL_ID_LENGTH).email_id


def test_rejects_an_id_beyond_the_maximum_length() -> None:
    with pytest.raises(InvalidRequestError):
        GetEmailRequest(email_id="a" * (MAX_EMAIL_ID_LENGTH + 1))
