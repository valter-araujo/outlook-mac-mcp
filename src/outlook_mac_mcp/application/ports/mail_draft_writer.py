from typing import Protocol

from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft


class MailDraftWriter(Protocol):
    def create(self, new_draft: NewMailDraft) -> CreatedMailDraft:
        """Save `new_draft` in the mailbox's Drafts folder and return its handle.

        Saves only: nothing is sent, and no other message is read or changed.
        """
        ...
