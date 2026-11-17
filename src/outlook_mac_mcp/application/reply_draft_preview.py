from dataclasses import dataclass

from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft


@dataclass(frozen=True, slots=True)
class ReplyDraftPreview:
    """A validated reply waiting for the user's go-ahead.

    Warnings are kept apart from the summary so a client can show them prominently
    instead of leaving them buried at the end of one long line.
    """

    token: str
    summary: str
    warnings: tuple[str, ...]
    draft: NewReplyDraft
