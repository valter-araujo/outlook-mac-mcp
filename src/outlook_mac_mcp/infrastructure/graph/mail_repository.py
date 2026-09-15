from collections.abc import Mapping
from http import HTTPStatus
from typing import Any
from urllib.parse import quote

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.errors import EmailNotFoundError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.email_mapper import (
    MESSAGE_FIELDS,
    to_email,
    to_email_detail,
)
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError

UNREAD_FILTER = "isRead eq false"
NEWEST_FIRST_ORDER = "receivedDateTime desc"
DETAIL_FIELDS = (*MESSAGE_FIELDS, "body")
BODY_AS_TEXT_HEADER = {"Prefer": 'outlook.body-content-type="text"'}


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

    def get_by_id(self, email_id: str) -> EmailDetail:
        """The id is percent-encoded before it becomes a path segment.

        Unlike the folder, it is caller-supplied: Graph ids can contain characters that
        are significant in a URL, and an unencoded one could otherwise change the path.
        """
        try:
            payload = self._client.get(
                f"/me/messages/{quote(email_id, safe='')}",
                {"$select": ",".join(DETAIL_FIELDS)},
                BODY_AS_TEXT_HEADER,
            )
        except GraphRequestError as error:
            if error.status_code == HTTPStatus.NOT_FOUND:
                raise EmailNotFoundError(f"no email with id {email_id}") from error
            raise
        return to_email_detail(payload)


def _read_messages(payload: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    messages = payload.get("value")
    if not isinstance(messages, list):
        raise GraphResponseError("the message collection carried no value array")
    for message in messages:
        if not isinstance(message, dict):
            raise GraphResponseError("the message collection held something other than a message")
    return messages
