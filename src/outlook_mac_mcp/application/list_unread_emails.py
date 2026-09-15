from dataclasses import dataclass

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.folder_name import FolderName


@dataclass(frozen=True, slots=True)
class ListUnreadEmailsRequest:
    folder: FolderName = FolderName.INBOX
    limit: int = DEFAULT_LIMIT

    def __post_init__(self) -> None:
        ensure_limit_within_bounds(self.limit)


class ListUnreadEmails:
    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, request: ListUnreadEmailsRequest) -> tuple[Email, ...]:
        return self._mail_repository.list_unread(request.folder, request.limit)
