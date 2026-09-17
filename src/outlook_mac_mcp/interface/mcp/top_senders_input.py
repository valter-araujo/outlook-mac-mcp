from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.top_senders import DEFAULT_TOP_SENDERS, TopSendersRequest
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.interface.mcp.email_filters_input import (
    FilterFolder,
    IsRead,
    ReceivedAfter,
    ReceivedBefore,
)

SenderLimit = Annotated[
    int,
    Field(ge=MIN_LIMIT, le=MAX_LIMIT, description="How many senders to return, most first."),
]


class TopSendersInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    folder: FilterFolder = FolderSelection.INBOX
    is_read: IsRead = None
    received_after: ReceivedAfter = None
    received_before: ReceivedBefore = None
    limit: SenderLimit = DEFAULT_TOP_SENDERS

    def to_request(self, folder_ids: tuple[str, ...]) -> TopSendersRequest:
        return TopSendersRequest(
            folders=folder_ids,
            is_read=self.is_read,
            received_after=self.received_after,
            received_before=self.received_before,
            limit=self.limit,
        )
