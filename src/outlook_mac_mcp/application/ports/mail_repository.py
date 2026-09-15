from typing import Protocol

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.folder_name import FolderName


class MailRepository(Protocol):
    def list_unread(self, folder: FolderName, limit: int) -> tuple[Email, ...]:
        """Return up to `limit` unread emails from `folder`, newest first."""
        ...
