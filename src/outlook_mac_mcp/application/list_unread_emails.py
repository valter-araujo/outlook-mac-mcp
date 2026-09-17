from dataclasses import dataclass

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.folder_selection import ensure_well_formed_folders
from outlook_mac_mcp.domain.page import Page


@dataclass(frozen=True, slots=True)
class ListUnreadEmailsRequest:
    folders: tuple[FolderName, ...] = (FolderName.INBOX,)
    limit: int = DEFAULT_LIMIT

    def __post_init__(self) -> None:
        ensure_limit_within_bounds(self.limit)
        ensure_well_formed_folders(self.folders)


class ListUnreadEmails:
    def __init__(self, mail_repository: MailRepository) -> None:
        self._mail_repository = mail_repository

    def execute(self, request: ListUnreadEmailsRequest) -> Page[Email]:
        return self._mail_repository.list_unread(request.folders, request.limit)
