from typing import Annotated

from pydantic import AwareDatetime, BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email_filters import (
    MAX_SENDER_ADDRESS_LENGTH,
    EmailFilters,
)
from outlook_mac_mcp.domain.folder_name import FolderName

# Mirrors SENDER_ADDRESS_PATTERN so the refusal appears in the tool schema.
SENDER_PATTERN = r"^[^@\s']+@[^@\s']+\.[^@\s']+$"

FilterFolder = Annotated[FolderName, Field(description="Mailbox folder to read.")]
IsRead = Annotated[
    bool | None,
    Field(description="true for read emails only, false for unread only, omit for both."),
]
Sender = Annotated[
    str | None,
    Field(
        pattern=SENDER_PATTERN,
        max_length=MAX_SENDER_ADDRESS_LENGTH,
        description="Exact sender address to match, e.g. ana@example.com. Omit for any sender.",
    ),
]
ReceivedAfter = Annotated[
    AwareDatetime | None,
    Field(description="Inclusive lower bound on the received time, ISO 8601 with a UTC offset."),
]
ReceivedBefore = Annotated[
    AwareDatetime | None,
    Field(description="Exclusive upper bound on the received time, ISO 8601 with a UTC offset."),
]
HasAttachments = Annotated[
    bool | None,
    Field(description="true for emails with attachments only, false for none, omit for both."),
]


class EmailFiltersInput(BaseModel):
    """The filter arguments shared by the listing and the count, turned into EmailFilters."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    folder: FilterFolder = FolderName.INBOX
    is_read: IsRead = None
    sender: Sender = None
    received_after: ReceivedAfter = None
    received_before: ReceivedBefore = None
    has_attachments: HasAttachments = None

    def to_filters(self) -> EmailFilters:
        return EmailFilters(
            folder=self.folder,
            is_read=self.is_read,
            sender=self.sender,
            received_after=self.received_after,
            received_before=self.received_before,
            has_attachments=self.has_attachments,
        )
