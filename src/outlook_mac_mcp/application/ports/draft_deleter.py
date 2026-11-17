from typing import Protocol

from outlook_mac_mcp.domain.saved_draft import SavedDraft


class DraftDeleter(Protocol):
    def get_draft(self, email_id: str) -> SavedDraft:
        """Return the message with `email_id`, only if it is an unsent draft in Drafts.

        Raises EmailNotFoundError when there is no such message, and NotADraftError when
        it exists anywhere else or has been sent.
        """
        ...

    def delete(self, email_id: str) -> None:
        """Move the message with `email_id` to Deleted Items, where it stays recoverable.

        Raises EmailNotFoundError when there is no such message. Callers check
        get_draft first; this method does not repeat the check.
        """
        ...
