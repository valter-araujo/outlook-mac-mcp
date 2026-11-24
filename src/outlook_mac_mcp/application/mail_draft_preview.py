from dataclasses import dataclass

from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft


@dataclass(frozen=True, slots=True)
class MailDraftPreview:
    """A validated draft waiting for the user's go-ahead.

    The token is the only handle a client gets back, so a create call can refer to
    exactly what was previewed and nothing else.
    """

    token: str
    summary: str
    warnings: tuple[str, ...]
    new_draft: NewMailDraft
