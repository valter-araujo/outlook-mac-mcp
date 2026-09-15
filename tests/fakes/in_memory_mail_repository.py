from collections import defaultdict

from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.errors import EmailNotFoundError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.domain.sort_order import SortOrder


class InMemoryMailRepository:
    """Fake MailRepository backed by a dict; sorts newest first like Graph does.

    Everything is in memory, so every total is exact.
    """

    def __init__(self) -> None:
        self._emails: dict[FolderName, list[Email]] = defaultdict(list)
        self._bodies: dict[str, str] = {}

    def add(self, folder: FolderName, email: Email, body: str = "") -> None:
        self._emails[folder].append(email)
        self._bodies[email.id] = body

    def list_unread(self, folder: FolderName, limit: int) -> Page[Email]:
        unread = [email for email in self._emails[folder] if not email.is_read]
        newest_first = sorted(unread, key=lambda email: email.received_at, reverse=True)
        return _exact_page(newest_first, limit)

    def search(self, request: SearchEmailsRequest) -> Page[Email]:
        """Insertion order, deliberately not date order: the port promises relevance."""
        needle = request.term.casefold()
        matches = [
            email
            for email in self._emails[request.folder]
            if _matches(email, needle, request.scope)
        ]
        return _exact_page(matches, request.limit)

    def list_matching(self, request: ListEmailsRequest) -> Page[Email]:
        matching = self._matching(request.filters)
        by_date = sorted(
            matching,
            key=lambda email: email.received_at,
            reverse=request.sort is SortOrder.NEWEST,
        )
        return _exact_page(by_date, request.limit)

    def count_matching(self, filters: EmailFilters) -> int:
        return len(self._matching(filters))

    def _matching(self, filters: EmailFilters) -> list[Email]:
        return [email for email in self._emails[filters.folder] if _passes(email, filters)]

    def get_by_id(self, email_id: str) -> EmailDetail:
        for emails in self._emails.values():
            for email in emails:
                if email.id == email_id:
                    return EmailDetail(email=email, body=self._bodies[email_id])
        raise EmailNotFoundError(f"no email with id {email_id}")


def _exact_page(emails: list[Email], limit: int) -> Page[Email]:
    return Page(items=tuple(emails[:limit]), total=len(emails), total_is_exact=True)


def _passes(email: Email, filters: EmailFilters) -> bool:
    """The same semantics the Graph query encodes: after inclusive, before exclusive."""
    if filters.is_read is not None and email.is_read is not filters.is_read:
        return False
    if filters.has_attachments is not None and email.has_attachments is not filters.has_attachments:
        return False
    if filters.sender is not None and email.sender.address.casefold() != filters.sender.casefold():
        return False
    if filters.received_after is not None and email.received_at < filters.received_after:
        return False
    return filters.received_before is None or email.received_at < filters.received_before


def _matches(email: Email, needle: str, scope: SearchScope) -> bool:
    if scope is SearchScope.SUBJECT:
        return needle in email.subject.casefold()
    if scope is SearchScope.SENDER:
        return (
            needle in email.sender.address.casefold()
            or needle in email.sender.display_name.casefold()
        )
    return (
        needle in email.subject.casefold()
        or needle in email.preview.casefold()
        or needle in email.sender.address.casefold()
    )
