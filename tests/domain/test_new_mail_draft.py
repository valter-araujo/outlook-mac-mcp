import pytest

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.new_mail_draft import (
    MAX_BODY_LENGTH,
    MAX_RECIPIENTS,
    MAX_SUBJECT_LENGTH,
    NewMailDraft,
)

ANA = EmailAddress(address="ana@example.com")


def recipients(count: int) -> tuple[EmailAddress, ...]:
    return tuple(EmailAddress(address=f"person{index}@example.com") for index in range(count))


def test_accepts_a_minimal_draft_with_one_recipient() -> None:
    draft = NewMailDraft(subject="Hello", body="", to=(ANA,))

    assert draft.cc == ()


@pytest.mark.parametrize("subject", ["", "a" * (MAX_SUBJECT_LENGTH + 1)])
def test_rejects_a_subject_outside_the_length_bounds(subject: str) -> None:
    with pytest.raises(InvalidRequestError, match="subject"):
        NewMailDraft(subject=subject, body="", to=(ANA,))


def test_rejects_a_body_beyond_the_maximum_length() -> None:
    with pytest.raises(InvalidRequestError, match="body"):
        NewMailDraft(subject="Hello", body="a" * (MAX_BODY_LENGTH + 1), to=(ANA,))


def test_accepts_a_body_at_exactly_the_maximum_length() -> None:
    draft = NewMailDraft(subject="Hello", body="a" * MAX_BODY_LENGTH, to=(ANA,))

    assert len(draft.body) == MAX_BODY_LENGTH


def test_rejects_a_draft_with_no_recipient_in_to() -> None:
    with pytest.raises(InvalidRequestError, match="recipient"):
        NewMailDraft(subject="Hello", body="", to=(), cc=(ANA,))


def test_rejects_more_recipients_than_the_limit_across_to_and_cc() -> None:
    with pytest.raises(InvalidRequestError, match=str(MAX_RECIPIENTS)):
        NewMailDraft(subject="Hello", body="", to=recipients(MAX_RECIPIENTS), cc=(ANA,))


def test_accepts_exactly_the_recipient_limit_across_to_and_cc() -> None:
    draft = NewMailDraft(subject="Hello", body="", to=recipients(30), cc=recipients(20))

    assert len(draft.to) + len(draft.cc) == MAX_RECIPIENTS
