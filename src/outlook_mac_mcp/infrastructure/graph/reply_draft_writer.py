from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft
from outlook_mac_mcp.domain.errors import (
    MailDraftWriteUnconfirmedError,
)
from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import address_key
from outlook_mac_mcp.domain.reply_source import ReplySource
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import optional_text, required_text
from outlook_mac_mcp.infrastructure.graph.message_path import (
    message_path,
    raise_message_lookup_error,
)
from outlook_mac_mcp.infrastructure.graph.reply_draft_payload import to_reply_draft_payload
from outlook_mac_mcp.infrastructure.graph.reply_source_mapper import (
    REPLY_SOURCE_FIELDS,
    to_reply_source,
)

ME_PATH = "/me"
OWN_ADDRESS_FIELDS = ("mail", "userPrincipalName")
REPLY_ACTIONS = {ReplyMode.REPLY: "createReply", ReplyMode.REPLY_ALL: "createReplyAll"}


class GraphReplyDraftWriter:
    """Saves replies to existing messages as drafts in their conversation.

    Satisfies the ReplyDraftWriter port structurally, and is wired only behind the mail
    draft flag, alongside GraphMailDraftWriter. The one write it makes is createReply or
    createReplyAll, which save an unsent draft; neither sends, and nothing here calls a
    send endpoint.
    """

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def get_reply_source(self, email_id: str) -> ReplySource:
        try:
            payload = self._client.get(
                message_path(email_id), {"$select": ",".join(REPLY_SOURCE_FIELDS)}
            )
        except GraphRequestError as error:
            raise_message_lookup_error(error, email_id)
        return to_reply_source(payload)

    def get_own_addresses(self) -> frozenset[str]:
        """Both fields, because a personal account may fill only one of them."""
        payload = self._client.get(ME_PATH, {"$select": ",".join(OWN_ADDRESS_FIELDS)})
        addresses = (optional_text(payload, field) for field in OWN_ADDRESS_FIELDS)
        return frozenset(address_key(address) for address in addresses if address)

    def create_reply(self, draft: NewReplyDraft) -> CreatedMailDraft:
        path = f"{message_path(draft.source_id)}/{REPLY_ACTIONS[draft.mode]}"
        try:
            payload = self._client.post(path, to_reply_draft_payload(draft))
        except GraphRequestError as error:
            raise_message_lookup_error(error, draft.source_id)
        try:
            return CreatedMailDraft(id=required_text(payload, "id"))
        except GraphResponseError as error:
            raise MailDraftWriteUnconfirmedError(
                f"the reply was saved, but its id could not be read back afterward: {error}. "
                "Check your Drafts folder directly before retrying."
            ) from error
