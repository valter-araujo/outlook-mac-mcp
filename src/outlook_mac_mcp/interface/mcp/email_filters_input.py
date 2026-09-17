from typing import Annotated

from pydantic import AwareDatetime, BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email_filters import (
    MAX_SENDER_ADDRESS_LENGTH,
    EmailFilters,
)
from outlook_mac_mcp.domain.folder_selection import FolderSelection

# Mirrors SENDER_ADDRESS_PATTERN so the refusal appears in the tool schema.
SENDER_PATTERN = r"^[^@\s']+@[^@\s']+\.[^@\s']+$"
MAX_FOLDER_ARGUMENT_LENGTH = 500

FilterFolder = Annotated[
    str,
    Field(
        min_length=1,
        max_length=MAX_FOLDER_ARGUMENT_LENGTH,
        description=(
            "Mailbox folder to read: a well-known name (default inbox), all (the five "
            "well-known folders only -- never custom folders), or a custom folder's "
            "full path as shown in list_folders' custom list, e.g. Entrevistas/Work/AWS."
        ),
    ),
]
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

    folder: FilterFolder = FolderSelection.INBOX
    is_read: IsRead = None
    sender: Sender = None
    received_after: ReceivedAfter = None
    received_before: ReceivedBefore = None
    has_attachments: HasAttachments = None

    def to_filters(self, folder_ids: tuple[str, ...]) -> EmailFilters:
        return EmailFilters(
            folders=folder_ids,
            is_read=self.is_read,
            sender=self.sender,
            received_after=self.received_after,
            received_before=self.received_before,
            has_attachments=self.has_attachments,
        )
