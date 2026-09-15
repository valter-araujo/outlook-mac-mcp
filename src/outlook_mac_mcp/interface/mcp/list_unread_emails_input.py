from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.list_unread_emails import (
    DEFAULT_LIMIT,
    MAX_LIMIT,
    MIN_LIMIT,
    ListUnreadEmailsRequest,
)
from outlook_mac_mcp.domain.folder_name import FolderName

Folder = Annotated[FolderName, Field(description="Mailbox folder to read.")]
Limit = Annotated[
    int,
    Field(ge=MIN_LIMIT, le=MAX_LIMIT, description="Maximum number of emails to return."),
]


class ListUnreadEmailsInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request.

    The bounds live in the `Limit` alias so the tool signature and this model advertise
    one rule from one place, and both derive it from the use case's own constants.
    """

    model_config = ConfigDict(frozen=True, extra="forbid")

    folder: Folder = FolderName.INBOX
    limit: Limit = DEFAULT_LIMIT

    def to_request(self) -> ListUnreadEmailsRequest:
        return ListUnreadEmailsRequest(folder=self.folder, limit=self.limit)
