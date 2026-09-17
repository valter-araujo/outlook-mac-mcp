from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_largest_emails import (
    DEFAULT_LARGEST_EMAILS,
    ListLargestEmailsRequest,
)
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.interface.mcp.email_filters_input import FilterFolder, IsRead, ReceivedAfter

LargestEmailsLimit = Annotated[
    int,
    Field(ge=MIN_LIMIT, le=MAX_LIMIT, description="How many emails to return, largest first."),
]


class ListLargestEmailsInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    folder: FilterFolder = FolderSelection.INBOX
    is_read: IsRead = None
    received_after: ReceivedAfter = None
    limit: LargestEmailsLimit = DEFAULT_LARGEST_EMAILS

    def to_request(self) -> ListLargestEmailsRequest:
        return ListLargestEmailsRequest(
            folders=self.folder.to_folders(),
            is_read=self.is_read,
            received_after=self.received_after,
            limit=self.limit,
        )
