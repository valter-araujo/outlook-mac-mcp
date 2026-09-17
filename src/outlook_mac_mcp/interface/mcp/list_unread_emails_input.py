from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmailsRequest
from outlook_mac_mcp.domain.folder_selection import FolderSelection

Folder = Annotated[
    FolderSelection,
    Field(
        description=(
            "Mailbox folder to read, or all to search every well-known folder and merge "
            "the results."
        )
    ),
]
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

    folder: Folder = FolderSelection.INBOX
    limit: Limit = DEFAULT_LIMIT

    def to_request(self) -> ListUnreadEmailsRequest:
        return ListUnreadEmailsRequest(folders=self.folder.to_folders(), limit=self.limit)
