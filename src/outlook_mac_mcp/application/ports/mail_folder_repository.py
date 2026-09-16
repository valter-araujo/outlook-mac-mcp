from typing import Protocol

from outlook_mac_mcp.domain.mail_folder import MailFolder


class MailFolderRepository(Protocol):
    def list_all(self) -> tuple[MailFolder, ...]:
        """Return every well-known folder, in FolderName declaration order.

        Arbitrary user-created folders are out of scope for v1 and never appear here.
        """
        ...
