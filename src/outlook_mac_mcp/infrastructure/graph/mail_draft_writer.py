from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft
from outlook_mac_mcp.domain.errors import MailDraftWriteUnconfirmedError
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import required_text
from outlook_mac_mcp.infrastructure.graph.mail_draft_payload import to_mail_draft_payload

MESSAGES_PATH = "/me/messages"


class GraphMailDraftWriter:
    """Saves drafts to the Drafts folder; the only code path that writes mail to Graph.

    Satisfies the MailDraftWriter port structurally. It is wired only when the draft
    flag is on, which is also the only time the token carries Mail.ReadWrite. The one
    call it makes is POST /me/messages, which creates an unsent message: there is no
    /send or /sendMail call anywhere in this project, so sending stays a manual step.
    """

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def create(self, new_draft: NewMailDraft) -> CreatedMailDraft:
        payload = self._client.post(MESSAGES_PATH, to_mail_draft_payload(new_draft))
        try:
            return CreatedMailDraft(id=required_text(payload, "id"))
        except GraphResponseError as error:
            raise MailDraftWriteUnconfirmedError(
                f"the draft was saved, but its id could not be read back afterward: {error}. "
                "Check your Drafts folder directly before retrying."
            ) from error
