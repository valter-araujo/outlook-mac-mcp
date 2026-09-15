from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.email_mapper import MESSAGE_FIELDS, to_email
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError

UNREAD_FILTER = "isRead eq false"
NEWEST_FIRST_ORDER = "receivedDateTime desc"


class GraphMailRepository:
    """Reads mail from Microsoft Graph for the signed-in account.

    Satisfies the MailRepository port structurally; the use cases never import this module.

    `$top` is a page size, not a cap: with a `$filter` Graph may return fewer items than
    asked while more unread mail exists behind `@odata.nextLink`. v1 reads a single page,
    so `limit` is an upper bound on what comes back, not a promise of what is there.
    """

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def list_unread(self, folder: FolderName, limit: int) -> tuple[Email, ...]:
        """Nothing user-supplied is spliced into the query: the folder segment comes from a
        closed enum of well-known names, the filter and order are constants, and `limit` is an
        int validated by the use case, so there is no string for a caller to break out of.
        """
        payload = self._client.get(
            f"/me/mailFolders/{folder.value}/messages",
            {
                "$filter": UNREAD_FILTER,
                "$orderby": NEWEST_FIRST_ORDER,
                "$top": limit,
                "$select": ",".join(MESSAGE_FIELDS),
            },
        )
        return tuple(to_email(message) for message in _read_messages(payload))


def _read_messages(payload: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    messages = payload.get("value")
    if not isinstance(messages, list):
        raise GraphResponseError("the message collection carried no value array")
    for message in messages:
        if not isinstance(message, dict):
            raise GraphResponseError("the message collection held something other than a message")
    return messages
