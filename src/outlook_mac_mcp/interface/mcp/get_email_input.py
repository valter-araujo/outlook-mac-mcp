from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.get_email import MAX_EMAIL_ID_LENGTH, GetEmailRequest

EmailId = Annotated[
    str,
    Field(
        min_length=1,
        max_length=MAX_EMAIL_ID_LENGTH,
        description="Graph id of the message, as returned by list_unread_emails.",
    ),
]


class GetEmailInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    email_id: EmailId

    def to_request(self) -> GetEmailRequest:
        return GetEmailRequest(email_id=self.email_id)
