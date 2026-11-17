import pytest

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.new_mail_draft import MAX_BODY_LENGTH, MAX_RECIPIENTS
from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import ReplyRecipients

ONE_RECIPIENT = ReplyRecipients(to=(EmailAddress(address="ana@example.com"),), cc=())


def test_accepts_an_empty_body() -> None:
    draft = NewReplyDraft(
        source_id="AAMk1", mode=ReplyMode.REPLY, body="", recipients=ONE_RECIPIENT
    )

    assert draft.body == ""


def test_rejects_a_reply_with_nobody_on_the_to_line() -> None:
    with pytest.raises(InvalidRequestError, match="extra_to"):
        NewReplyDraft(
            source_id="AAMk1",
            mode=ReplyMode.REPLY,
            body="Hi",
            recipients=ReplyRecipients(to=(), cc=()),
        )


def test_rejects_a_body_over_the_limit() -> None:
    with pytest.raises(InvalidRequestError, match="body"):
        NewReplyDraft(
            source_id="AAMk1",
            mode=ReplyMode.REPLY,
            body="x" * (MAX_BODY_LENGTH + 1),
            recipients=ONE_RECIPIENT,
        )


def test_rejects_more_recipients_than_the_limit() -> None:
    many = tuple(EmailAddress(address=f"p{index}@example.com") for index in range(MAX_RECIPIENTS))
    recipients = ReplyRecipients(to=many, cc=(EmailAddress(address="one@example.com"),))

    with pytest.raises(InvalidRequestError, match="recipients"):
        NewReplyDraft(source_id="AAMk1", mode=ReplyMode.REPLY_ALL, body="Hi", recipients=recipients)
