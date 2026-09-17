from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder


class CustomMailFolderView(BaseModel):
    """JSON-serializable projection of a CustomMailFolder."""

    model_config = ConfigDict(frozen=True)

    folder_id: str = Field(description="Graph's own id for this folder.")
    display_name: str = Field(description="The folder's own label; not unique across the tree.")
    path: str = Field(
        description=(
            "The full path from the mailbox root or a well-known folder, e.g. "
            "'Candidaturas/2026'. Two folders can share a display_name; path is what "
            "tells them apart."
        )
    )
    unread_count: int
    total_count: int

    @classmethod
    def from_folder(cls, folder: CustomMailFolder) -> "CustomMailFolderView":
        return cls(
            folder_id=folder.folder_id,
            display_name=folder.display_name,
            path=folder.path,
            unread_count=folder.unread_count,
            total_count=folder.total_count,
        )
