from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.mail_folder import MailFolder
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.custom_folder_scan import scan_custom_folders
from outlook_mac_mcp.infrastructure.graph.mail_folder_mapper import to_mail_folder

FOLDER_FIELDS = ("displayName", "unreadItemCount", "totalItemCount")


class GraphMailFolderRepository:
    """Reads each well-known folder from Microsoft Graph, one request per folder.

    Graph's mailFolder resource carries no well-known-name flag, so the collection
    endpoint cannot tell "Inbox" apart from a user-created folder without a locale-
    dependent guess at the display name. Fetching each well-known folder by its own
    well-known path segment avoids that guess entirely, at a fixed cost of five requests.

    Satisfies the MailFolderRepository port structurally.
    """

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def list_all(self) -> tuple[MailFolder, ...]:
        return tuple(self._read_folder(name) for name in FolderName)

    def list_custom(self, max_depth: int, max_folders: int) -> CustomFolderScan:
        return scan_custom_folders(self._client, max_depth, max_folders)

    def _read_folder(self, name: FolderName) -> MailFolder:
        payload = self._client.get(
            f"/me/mailFolders/{name.value}", {"$select": ",".join(FOLDER_FIELDS)}
        )
        return to_mail_folder(name, payload)
