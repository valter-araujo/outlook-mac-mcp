from outlook_mac_mcp.application.list_custom_folders import MAX_FOLDER_DEPTH, MAX_FOLDERS_SCANNED
from outlook_mac_mcp.application.ports.mail_folder_repository import MailFolderRepository
from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder
from outlook_mac_mcp.domain.errors import CustomFolderNotFoundError
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.domain.resolved_folder import ResolvedFolder


class ResolveFolders:
    """Turns a mail tool's raw `folder` argument -- a well-known name, `all`, or a
    custom folder's full path -- into the folders to actually query.

    A custom path is resolved fresh on every call, by walking the same tree
    list_folders walks (MailFolderRepository.list_custom, sharing its depth/folder-count
    caps by default): nothing here is cached, so a custom folder argument costs one
    full tree walk each time it's used.
    """

    def __init__(
        self,
        mail_folder_repository: MailFolderRepository,
        max_depth: int = MAX_FOLDER_DEPTH,
        max_folders: int = MAX_FOLDERS_SCANNED,
    ) -> None:
        self._mail_folder_repository = mail_folder_repository
        self._max_depth = max_depth
        self._max_folders = max_folders

    def execute(self, raw: str) -> ResolvedFolder:
        try:
            selection = FolderSelection(raw)
        except ValueError:
            return self._resolve_custom_path(raw)
        return ResolvedFolder(folder_ids=selection.to_folders(), echo=selection.describe_folders())

    def _resolve_custom_path(self, path: str) -> ResolvedFolder:
        scan = self._mail_folder_repository.list_custom(self._max_depth, self._max_folders)
        match = _find_by_path(scan, path)
        if match is None:
            raise CustomFolderNotFoundError(
                f"no custom folder with path {path!r} found within {self._max_depth} "
                f"levels deep and {self._max_folders} folders scanned"
            )
        return ResolvedFolder(folder_ids=(match.folder_id,), echo=match.path)


def _find_by_path(scan: CustomFolderScan, path: str) -> CustomMailFolder | None:
    return next((folder for folder in scan.folders if folder.path == path), None)
