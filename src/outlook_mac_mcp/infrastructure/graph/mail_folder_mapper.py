from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.mail_folder import MailFolder
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import optional_text


def to_mail_folder(well_known_name: FolderName, folder: Mapping[str, Any]) -> MailFolder:
    """The well-known name is the caller's, not read from the resource: Graph's response
    carries no such flag, which is why each folder is fetched by its own well-known path.
    """
    return MailFolder(
        well_known_name=well_known_name,
        display_name=optional_text(folder, "displayName"),
        unread_count=_required_count(folder, "unreadItemCount"),
        total_count=_required_count(folder, "totalItemCount"),
    )


def _required_count(folder: Mapping[str, Any], field: str) -> int:
    value = folder.get(field)
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise GraphResponseError(f"the folder carried no {field}")
    return value
