from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.mail_draft_preview import MailDraftPreview
from outlook_mac_mcp.application.ports.mail_draft_writer import MailDraftWriter
from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft


class CreateMailDraft:
    """Second half of a draft: the only step that changes the mailbox, and it only saves.

    The preview is taken before the write, not after: a retry following a failure past
    the wire must not be able to save the same draft twice, at the cost of previewing
    again.
    """

    def __init__(
        self, draft_store: DraftStore[MailDraftPreview], mail_draft_writer: MailDraftWriter
    ) -> None:
        self._draft_store = draft_store
        self._mail_draft_writer = mail_draft_writer

    def execute(self, token: str) -> CreatedMailDraft:
        preview = self._draft_store.take(token)
        return self._mail_draft_writer.create(preview.new_draft)
