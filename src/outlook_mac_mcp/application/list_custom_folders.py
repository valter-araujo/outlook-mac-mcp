from outlook_mac_mcp.application.ports.mail_folder_repository import MailFolderRepository
from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan

# Ten levels of nesting comfortably covers any mailbox a person organizes by hand.
MAX_FOLDER_DEPTH = 10
# A generous ceiling on how many folders one call may walk, so a pathological mailbox
# cannot turn one tool call into thousands of requests.
MAX_FOLDERS_SCANNED = 200


class ListCustomFolders:
    """Discovers every user-created folder, however deep, up to the caps below.

    Both caps are constructor arguments so a test can drive the partial-coverage paths
    without a mailbox that actually holds hundreds of folders or ten levels of nesting.
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

    def execute(self) -> CustomFolderScan:
        return self._mail_folder_repository.list_custom(self._max_depth, self._max_folders)
