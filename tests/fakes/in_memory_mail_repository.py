from collections import defaultdict

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.errors import EmailNotFoundError
from outlook_mac_mcp.domain.folder_name import FolderName


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

    def get_by_id(self, email_id: str) -> EmailDetail:
        for emails in self._emails.values():
            for email in emails:
                if email.id == email_id:
                    return EmailDetail(email=email, body=self._bodies[email_id])
        raise EmailNotFoundError(f"no email with id {email_id}")
