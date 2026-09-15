from dataclasses import dataclass

from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName

MIN_LIMIT = 1
MAX_LIMIT = 100


@dataclass(frozen=True, slots=True)
class ListUnreadEmailsRequest:
    folder: FolderName = FolderName.INBOX
    limit: int = 20

    def __post_init__(self) -> None:
        if not MIN_LIMIT <= self.limit <= MAX_LIMIT:
            raise InvalidRequestError(f"limit must be between {MIN_LIMIT} and {MAX_LIMIT}")


class ListUnreadEmails:
    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, request: ListUnreadEmailsRequest) -> tuple[Email, ...]:
        return self._mail_repository.list_unread(request.folder, request.limit)
