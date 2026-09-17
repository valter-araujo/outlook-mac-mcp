from typing import Protocol

from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.mail_folder import MailFolder


class MailFolderRepository(Protocol):
    def list_all(self) -> tuple[MailFolder, ...]:
        """Return every well-known folder, in FolderName declaration order.

        Arbitrary user-created folders are out of scope here and never appear in the
        result; see `list_custom` for those.
        """
        ...

    def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
        """Walk every user-created folder in the mailbox tree, at most `max_depth`
        levels deep and `max_folders` folders in all across the whole walk.
        """
        ...
