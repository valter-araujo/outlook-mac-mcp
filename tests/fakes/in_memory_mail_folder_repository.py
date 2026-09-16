from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.mail_folder import MailFolder


class InMemoryMailFolderRepository:
    """Fake MailFolderRepository; returns only the folders explicitly added, in
    FolderName declaration order, the way the real adapter always returns all five.
    """

    def __init__(self) -> None:
        self._folders: dict[FolderName, MailFolder] = {}

    def add(self, folder: MailFolder) -> None:
        self._folders[folder.well_known_name] = folder

    def list_all(self) -> tuple[MailFolder, ...]:
        return tuple(self._folders[name] for name in FolderName if name in self._folders)
