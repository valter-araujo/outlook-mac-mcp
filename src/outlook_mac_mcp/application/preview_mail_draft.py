from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.mail_draft_preview import MailDraftPreview
from outlook_mac_mcp.application.mail_draft_summary import describe_mail_draft
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft


class PreviewMailDraft:
    """First half of a draft: nothing leaves the process here.

    The draft is already validated by construction; this step only parks it under a
    token and describes it, so the user confirms exactly what CreateMailDraft will save.
    """

    def __init__(self, draft_store: DraftStore[MailDraftPreview]) -> None:
        self._draft_store = draft_store

    def execute(self, new_draft: NewMailDraft) -> MailDraftPreview:
        return self._draft_store.add(
            lambda token: MailDraftPreview(
                token=token, summary=describe_mail_draft(new_draft), new_draft=new_draft
            )
        )
