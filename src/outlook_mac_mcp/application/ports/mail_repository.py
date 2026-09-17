from typing import Protocol

from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.email_search_page import EmailSearchPage
from outlook_mac_mcp.domain.email_size_scan import EmailSizeScan
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.sender_scan import SenderScan


class MailRepository(Protocol):
    def list_unread(self, folders: tuple[str, ...], limit: int) -> Page[Email]:
        """Return up to `limit` unread emails from `folders`, newest first.

        More than one folder means every matching email is merged into one newest-first
        list before `limit` is applied, not `limit` per folder. The page's total is the
        number of unread emails across every folder given, exact regardless of how many.
        """
        ...

    def get_by_id(self, email_id: str) -> EmailDetail:
        """Return the email with `email_id`, body included.

        Raises EmailNotFoundError when the mailbox holds no such message.
        """
        ...

    def search(self, request: SearchEmailsRequest) -> EmailSearchPage:
        """Return the emails matching `request`, in relevance order.

        The term is matched as literal text; it never carries query operators. More than
        one folder in `request.folders` means each is searched and the results merged,
        grouped by folder rather than globally re-ranked (there is no cross-folder
        relevance score to merge by). The page's total is the number of matches across
        every folder given, which the backend may only be able to bound from below;
        `total_is_exact` says which. When `request.page_token` is given, that
        continuation token is followed directly instead of the query being rebuilt from
        `term`/`folders`/`scope`.
        """
        ...

    def list_matching(self, request: ListEmailsRequest) -> Page[Email]:
        """Return up to `request.limit` emails matching `request.filters`, in date order.

        More than one folder in `request.filters.folders` means every matching email is
        merged into one date-ordered list before `request.limit` is applied. The page's
        total is the number of matching emails across every folder given, exact.
        """
        ...

    def count_matching(self, filters: EmailFilters) -> int:
        """Return how many emails match `filters`, exactly, summed across every folder
        in `filters.folders`.
        """
        ...

    def scan_senders(self, filters: EmailFilters, ceiling: int) -> SenderScan:
        """Return the sender of each email matching `filters`, at most `ceiling` of them
        in all, however many folders `filters.folders` names.

        The scan's total is exact even when the walk stopped at the ceiling: an exact
        count is cheap per folder regardless of how much of that folder the walk itself
        goes on to cover.
        """
        ...

    def scan_email_sizes(self, filters: EmailFilters, ceiling: int) -> EmailSizeScan:
        """Return the size of each email matching `filters`, at most `ceiling` of them
        in all, however many folders `filters.folders` names.

        The scan's total is exact even when the walk stopped at the ceiling, for the
        same reason as `scan_senders`.
        """
        ...
