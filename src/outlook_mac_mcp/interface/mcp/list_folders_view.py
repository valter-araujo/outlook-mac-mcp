from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.mail_folder import MailFolder
from outlook_mac_mcp.interface.mcp.custom_mail_folder_view import CustomMailFolderView
from outlook_mac_mcp.interface.mcp.mail_folder_view import MailFolderView


class ListFoldersView(BaseModel):
    """The well-known folders and the discovered custom folders, reported separately
    so a caller never mistakes one set for the other.
    """

    model_config = ConfigDict(frozen=True)

    well_known: list[MailFolderView] = Field(
        description="The five well-known folders (inbox, archive, junk email, sent items, "
        "drafts): always the complete set."
    )
    custom: list[CustomMailFolderView] = Field(
        description="User-created folders found by walking the mailbox tree, each identified "
        "by its full path since two can share a display name."
    )
    depth_limit_reached: bool = Field(
        description="Whether some branch of the folder tree went deeper than the scan's depth "
        "limit and was left unwalked beyond that point; custom may then be missing folders."
    )
    folder_limit_reached: bool = Field(
        description="Whether the scan stopped before visiting every folder because it reached "
        "its folder-count limit; custom may then be missing folders."
    )

    @classmethod
    def from_result(
        cls, well_known: tuple[MailFolder, ...], custom: CustomFolderScan
    ) -> "ListFoldersView":
        return cls(
            well_known=[MailFolderView.from_folder(folder) for folder in well_known],
            custom=[CustomMailFolderView.from_folder(folder) for folder in custom.folders],
            depth_limit_reached=custom.depth_limit_reached,
            folder_limit_reached=custom.folder_limit_reached,
        )
