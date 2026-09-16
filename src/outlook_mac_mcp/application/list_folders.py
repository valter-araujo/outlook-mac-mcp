from outlook_mac_mcp.application.ports.mail_folder_repository import MailFolderRepository
from outlook_mac_mcp.domain.mail_folder import MailFolder


class ListFolders:
    """Takes no arguments: the well-known set is fixed, so there is nothing to filter."""

    def __init__(self, mail_folder_repository: MailFolderRepository) -> None:
        self._mail_folder_repository = mail_folder_repository

    def execute(self) -> tuple[MailFolder, ...]:
        return self._mail_folder_repository.list_all()
