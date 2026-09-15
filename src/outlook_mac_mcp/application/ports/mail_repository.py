from typing import Protocol

from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.page import Page


class MailRepository(Protocol):
    def list_unread(self, folder: FolderName, limit: int) -> Page[Email]:
        """Return up to `limit` unread emails from `folder`, newest first.

        The page's total is the number of unread emails in the folder, exact.
        """
        ...

    def get_by_id(self, email_id: str) -> EmailDetail:
        """Return the email with `email_id`, body included.

        Raises EmailNotFoundError when the mailbox holds no such message.
        """
        ...

    def search(self, request: SearchEmailsRequest) -> Page[Email]:
        """Return the emails matching `request`, in relevance order.

        The term is matched as literal text; it never carries query operators. The
        page's total is the number of matches, which the backend may only be able to
        bound from below; `total_is_exact` says which.
        """
        ...
