from collections import defaultdict

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.folder_name import FolderName


class InMemoryMailRepository:
    """Fake MailRepository backed by a dict; sorts newest first like Graph does."""

    def __init__(self) -> None:
        self._emails: dict[FolderName, list[Email]] = defaultdict(list)

    def add(self, folder: FolderName, email: Email) -> None:
        self._emails[folder].append(email)

    def list_unread(self, folder: FolderName, limit: int) -> tuple[Email, ...]:
        unread = [email for email in self._emails[folder] if not email.is_read]
        newest_first = sorted(unread, key=lambda email: email.received_at, reverse=True)
        return tuple(newest_first[:limit])
