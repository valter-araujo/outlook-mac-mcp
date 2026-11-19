from dataclasses import dataclass

from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.new_mail_draft import MAX_BODY_LENGTH, MAX_RECIPIENTS
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import ReplyRecipients


@dataclass(frozen=True, slots=True)
class NewReplyDraft:
    """A reply as the user confirmed it, validated before anything is previewed.

    `body` is the new text only, plain text for the same reason NewMailDraft's is; the
    quoted original is added by the mailbox when the draft is saved and never passes
    through this project.
    """

    source_id: str
    mode: ReplyMode
    body: str
    recipients: ReplyRecipients

    def __post_init__(self) -> None:
        if len(self.body) > MAX_BODY_LENGTH:
            raise InvalidRequestError(f"body must be at most {MAX_BODY_LENGTH} characters")
        if not self.recipients.to:
            raise InvalidRequestError(
                "the reply has no recipient left once your own address is removed; "
                "name one with extra_to"
            )
        if len(self.recipients.to) + len(self.recipients.cc) > MAX_RECIPIENTS:
            raise InvalidRequestError(f"at most {MAX_RECIPIENTS} recipients are allowed")
