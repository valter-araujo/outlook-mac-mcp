from collections import defaultdict

from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.errors import EmailNotFoundError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.search_scope import SearchScope


class InMemoryMailRepository:
    """Fake MailRepository backed by a dict; sorts newest first like Graph does."""

    def __init__(self) -> None:
        self._emails: dict[FolderName, list[Email]] = defaultdict(list)
        self._bodies: dict[str, str] = {}

    def add(self, folder: FolderName, email: Email, body: str = "") -> None:
        self._emails[folder].append(email)
        self._bodies[email.id] = body

    def list_unread(self, folder: FolderName, limit: int) -> tuple[Email, ...]:
        unread = [email for email in self._emails[folder] if not email.is_read]
        newest_first = sorted(unread, key=lambda email: email.received_at, reverse=True)
        return tuple(newest_first[:limit])

    def search(self, request: SearchEmailsRequest) -> tuple[Email, ...]:
        """Insertion order, deliberately not date order: the port promises relevance."""
        needle = request.term.casefold()
        matches = [
            email
            for email in self._emails[request.folder]
            if _matches(email, needle, request.scope)
        ]
        return tuple(matches[: request.limit])

    def get_by_id(self, email_id: str) -> EmailDetail:
        for emails in self._emails.values():
            for email in emails:
                if email.id == email_id:
                    return EmailDetail(email=email, body=self._bodies[email_id])
        raise EmailNotFoundError(f"no email with id {email_id}")


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
