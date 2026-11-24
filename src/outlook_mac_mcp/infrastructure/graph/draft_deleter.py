from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.errors import (
    NotADraftError,
)
from outlook_mac_mcp.domain.saved_draft import SavedDraft
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import to_email_addresses
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
from outlook_mac_mcp.infrastructure.graph.json_fields import (
    optional_text,
    required_datetime,
    required_flag,
    required_text,
)
from outlook_mac_mcp.infrastructure.graph.message_path import (
    message_path,
    raise_message_lookup_error,
)

DRAFTS_FOLDER_PATH = "/me/mailFolders/drafts"
DELETED_ITEMS_FOLDER = "deleteditems"
DRAFT_FIELDS = (
    "id",
    "subject",
    "toRecipients",
    "lastModifiedDateTime",
    "isDraft",
    "parentFolderId",
)


class GraphDraftDeleter:
    """Deletes drafts, and only drafts that sit in the Drafts folder.

    Satisfies the DraftDeleter port structurally; wired only when both the mail draft
    flag and the draft delete flag are on. "Delete" is a move to Deleted Items, never
    an HTTP DELETE and never permanentDelete: a live check on 2026-11-19 found that
    Graph's DELETE on a personal Outlook.com account does not put the message in
    Deleted Items (it was found nowhere the user can see), so only an explicit move
    keeps a mistaken deletion visible and recoverable in Outlook.

    Both isDraft and the parent folder are checked: isDraft alone would accept an unsent
    message the user filed elsewhere, and the folder alone would accept a sent message
    someone moved into Drafts.
    """

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def get_draft(self, email_id: str) -> SavedDraft:
        message = self._get_message(email_id)
        if not required_flag(message, "isDraft") or not self._is_in_drafts(message):
            raise NotADraftError(
                f"message {email_id} is not an unsent draft in the Drafts folder; "
                "only those can be deleted here"
            )
        return SavedDraft(
            id=required_text(message, "id"),
            subject=optional_text(message, "subject"),
            to=to_email_addresses(message.get("toRecipients")),
            last_modified_at=required_datetime(message, "lastModifiedDateTime"),
        )

    def delete(self, email_id: str) -> None:
        try:
            self._client.post(
                f"{message_path(email_id)}/move", {"destinationId": DELETED_ITEMS_FOLDER}
            )
        except GraphRequestError as error:
            raise_message_lookup_error(error, email_id)

    def _get_message(self, email_id: str) -> Mapping[str, Any]:
        try:
            return self._client.get(message_path(email_id), {"$select": ",".join(DRAFT_FIELDS)})
        except GraphRequestError as error:
            raise_message_lookup_error(error, email_id)

    def _is_in_drafts(self, message: Mapping[str, Any]) -> bool:
        drafts = self._client.get(DRAFTS_FOLDER_PATH, {"$select": "id"})
        return required_text(message, "parentFolderId") == required_text(drafts, "id")
