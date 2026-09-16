from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.mail_folder import MailFolder


class MailFolderView(BaseModel):
    """JSON-serializable projection of a MailFolder."""

    model_config = ConfigDict(frozen=True)

    well_known_name: FolderName = Field(
        description="The well-known name, usable as the folder argument of every other tool."
    )
    display_name: str = Field(description="The mailbox's own, possibly localized, label.")
    unread_count: int
    total_count: int

    @classmethod
    def from_folder(cls, folder: MailFolder) -> "MailFolderView":
        return cls(
            well_known_name=folder.well_known_name,
            display_name=folder.display_name,
            unread_count=folder.unread_count,
            total_count=folder.total_count,
        )
