from typing import Protocol

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.folder_name import FolderName


class MailRepository(Protocol):
    def list_unread(self, folder: FolderName, limit: int) -> tuple[Email, ...]:
        """Return up to `limit` unread emails from `folder`, newest first."""
        ...

    def get_by_id(self, email_id: str) -> EmailDetail:
        """Return the email with `email_id`, body included.

        Raises EmailNotFoundError when the mailbox holds no such message.
        """
        ...

    def search(self, folder: FolderName, term: str, limit: int) -> tuple[Email, ...]:
        """Return up to `limit` emails in `folder` matching `term`, in relevance order.

        `term` is matched as literal text; it never carries query operators.
        """
        ...
