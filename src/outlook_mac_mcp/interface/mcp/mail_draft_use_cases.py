from dataclasses import dataclass

from outlook_mac_mcp.application.create_mail_draft import CreateMailDraft
from outlook_mac_mcp.application.create_reply_draft import CreateReplyDraft
from outlook_mac_mcp.application.preview_mail_draft import PreviewMailDraft
from outlook_mac_mcp.application.preview_reply_draft import PreviewReplyDraft


@dataclass(frozen=True, slots=True)
class MailDraftUseCases:
    """The draft pairs, bundled so the server registers them together behind one flag.

    Each pair has its own draft store, so a token from one pair, or from a calendar pair,
    is never accepted by another.
    """

    preview_mail_draft: PreviewMailDraft
    create_mail_draft: CreateMailDraft
    preview_reply_draft: PreviewReplyDraft
    create_reply_draft: CreateReplyDraft
