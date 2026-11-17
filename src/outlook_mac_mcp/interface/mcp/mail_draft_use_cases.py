from dataclasses import dataclass

from outlook_mac_mcp.application.create_mail_draft import CreateMailDraft
from outlook_mac_mcp.application.preview_mail_draft import PreviewMailDraft


@dataclass(frozen=True, slots=True)
class MailDraftUseCases:
    """The one preview/apply pair for drafts, bundled so the server registers both or
    neither. Its draft store is its own, so a token from a calendar pair is never
    accepted here.
    """

    preview_mail_draft: PreviewMailDraft
    create_mail_draft: CreateMailDraft
