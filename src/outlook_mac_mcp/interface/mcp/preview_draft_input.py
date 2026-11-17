from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_mail_draft import (
    MAX_BODY_LENGTH,
    MAX_RECIPIENTS,
    MAX_SUBJECT_LENGTH,
    MIN_SUBJECT_LENGTH,
    NewMailDraft,
)
from outlook_mac_mcp.interface.mcp.preview_event_input import EMAIL_PATTERN, MAX_ADDRESS_LENGTH

DraftSubject = Annotated[
    str,
    Field(
        min_length=MIN_SUBJECT_LENGTH,
        max_length=MAX_SUBJECT_LENGTH,
        description="Subject line of the draft.",
    ),
]
DraftBody = Annotated[
    str,
    Field(
        max_length=MAX_BODY_LENGTH,
        description=(
            "Plain text only, up to 32,768 characters, or empty. No HTML: a body can "
            "originate from untrusted content (an email or a web page) and markup there "
            "is the same injection surface get_email's body warning covers."
        ),
    ),
]
Recipient = Annotated[str, Field(pattern=EMAIL_PATTERN, max_length=MAX_ADDRESS_LENGTH)]
# Tuples so the tool signature can carry an immutable empty default.
ToRecipients = Annotated[
    tuple[Recipient, ...],
    Field(
        min_length=1,
        max_length=MAX_RECIPIENTS,
        description="Email addresses for the To line; at least one. To and cc together: 50.",
    ),
]
CcRecipients = Annotated[
    tuple[Recipient, ...],
    Field(max_length=MAX_RECIPIENTS, description="Email addresses for the Cc line, or none."),
]


class PreviewDraftInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a NewMailDraft."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    subject: DraftSubject
    body: DraftBody
    to: ToRecipients
    cc: CcRecipients = ()

    def to_new_mail_draft(self) -> NewMailDraft:
        return NewMailDraft(
            subject=self.subject,
            body=self.body,
            to=tuple(EmailAddress(address=address) for address in self.to),
            cc=tuple(EmailAddress(address=address) for address in self.cc),
        )
