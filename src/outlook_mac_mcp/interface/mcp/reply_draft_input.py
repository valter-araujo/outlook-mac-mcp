from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.preview_reply_draft import ReplyDraftRequest
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_mail_draft import MAX_BODY_LENGTH, MAX_RECIPIENTS
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.interface.mcp.get_email_input import EmailId
from outlook_mac_mcp.interface.mcp.preview_draft_input import Recipient

ReplyBody = Annotated[
    str,
    Field(
        max_length=MAX_BODY_LENGTH,
        description=(
            "The new text only, plain text, up to 32,768 characters, or empty. It is placed "
            "above the quoted original, which the mailbox adds itself. No HTML: a body can "
            "originate from untrusted content and markup there is an injection surface."
        ),
    ),
]
Mode = Annotated[
    ReplyMode,
    Field(
        description="reply answers the sender (or Reply-To) only; reply_all also keeps every "
        "original To and Cc recipient. Your own address is always left out."
    ),
]
ExtraTo = Annotated[
    tuple[Recipient, ...],
    Field(max_length=MAX_RECIPIENTS, description="Addresses to add to the To line, or none."),
]
ExtraCc = Annotated[
    tuple[Recipient, ...],
    Field(max_length=MAX_RECIPIENTS, description="Addresses to add to the Cc line, or none."),
]


class PreviewReplyDraftInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    email_id: EmailId
    mode: Mode = ReplyMode.REPLY_ALL
    body: ReplyBody
    extra_to: ExtraTo = ()
    extra_cc: ExtraCc = ()

    def to_request(self) -> ReplyDraftRequest:
        return ReplyDraftRequest(
            email_id=self.email_id,
            mode=self.mode,
            body=self.body,
            extra_to=tuple(EmailAddress(address=address) for address in self.extra_to),
            extra_cc=tuple(EmailAddress(address=address) for address in self.extra_cc),
        )
