from typing import Protocol

from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.email_search_page import EmailSearchPage
from outlook_mac_mcp.domain.email_size_scan import EmailSizeScan
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.sender_scan import SenderScan


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

    def search(self, request: SearchEmailsRequest) -> EmailSearchPage:
        """Return the emails matching `request`, in relevance order.

        The term is matched as literal text; it never carries query operators. The
        page's total is the number of matches, which the backend may only be able to
        bound from below; `total_is_exact` says which. When `request.page_token` is
        given, that continuation token is followed directly instead of the query being
        rebuilt from `term`/`folder`/`scope`.
        """
        ...

    def list_matching(self, request: ListEmailsRequest) -> Page[Email]:
        """Return up to `request.limit` emails matching `request.filters`, in date order.

        The page's total is the number of matching emails in the folder, exact.
        """
        ...

    def count_matching(self, filters: EmailFilters) -> int:
        """Return how many emails match `filters`, exactly."""
        ...

    def scan_senders(self, filters: EmailFilters, ceiling: int) -> SenderScan:
        """Return the sender of each email matching `filters`, at most `ceiling` of them.

        The scan's total is exact even when the walk stopped at the ceiling.
        """
        ...

    def scan_email_sizes(self, filters: EmailFilters, ceiling: int) -> EmailSizeScan:
        """Return the size of each email matching `filters`, at most `ceiling` of them.

        The scan's total is exact even when the walk stopped at the ceiling.
        """
        ...
