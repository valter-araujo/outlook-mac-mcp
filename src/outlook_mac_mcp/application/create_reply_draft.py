from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.ports.reply_draft_writer import ReplyDraftWriter
from outlook_mac_mcp.application.reply_draft_preview import ReplyDraftPreview
from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft


class CreateReplyDraft:
    """Second half of a reply: the only step that changes the mailbox, and it only saves.

    The preview is taken before the write for the same reason CreateMailDraft takes it
    first: a retry after a failure past the wire must not save the reply twice.
    """

    def __init__(
        self, draft_store: DraftStore[ReplyDraftPreview], reply_draft_writer: ReplyDraftWriter
    ) -> None:
        self._draft_store = draft_store
        self._reply_draft_writer = reply_draft_writer

    def execute(self, token: str) -> CreatedMailDraft:
        preview = self._draft_store.take(token)
        return self._reply_draft_writer.create_reply(preview.draft)
