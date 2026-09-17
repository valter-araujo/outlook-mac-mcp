from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.mail_folder import MailFolder

EMPTY_CUSTOM_SCAN = CustomFolderScan(
    folders=(), depth_limit_reached=False, folder_limit_reached=False
)


class InMemoryMailFolderRepository:
    """Fake MailFolderRepository; returns only the folders explicitly added, in
    FolderName declaration order, the way the real adapter always returns all five.

    `list_custom` returns whatever scan was set with `set_custom`, ignoring
    `max_depth`/`max_folders`: the walking algorithm those caps govern lives in the
    Graph adapter and is tested there against respx, not re-simulated here.
    """

    def __init__(self) -> None:
        self._folders: dict[FolderName, MailFolder] = {}
        self._custom: CustomFolderScan = EMPTY_CUSTOM_SCAN

    def add(self, folder: MailFolder) -> None:
        self._folders[folder.well_known_name] = folder

    def set_custom(self, scan: CustomFolderScan) -> None:
        self._custom = scan

    def list_all(self) -> tuple[MailFolder, ...]:
        return tuple(self._folders[name] for name in FolderName if name in self._folders)

    def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
        return self._custom
