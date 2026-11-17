from outlook_mac_mcp.domain.errors import EmailNotFoundError, NotADraftError
from outlook_mac_mcp.domain.saved_draft import SavedDraft


class InMemoryDraftDeleter:
    """Fake DraftDeleter: `drafts` are in Drafts, `other_ids` exist anywhere else."""

    def __init__(
        self, drafts: tuple[SavedDraft, ...] = (), other_ids: frozenset[str] = frozenset()
    ) -> None:
        self._drafts = {draft.id: draft for draft in drafts}
        self._other_ids = set(other_ids)
        self.deleted: list[str] = []

    def get_draft(self, email_id: str) -> SavedDraft:
        if email_id in self._other_ids:
            raise NotADraftError(f"message {email_id} is not a draft in Drafts")
        try:
            return self._drafts[email_id]
        except KeyError:
            raise EmailNotFoundError(f"no email with id {email_id}") from None

    def delete(self, email_id: str) -> None:
        self._drafts.pop(email_id)
        self.deleted.append(email_id)

    def move_out_of_drafts(self, email_id: str) -> None:
        """Simulate the user sending or filing the draft after it was previewed."""
        self._drafts.pop(email_id)
        self._other_ids.add(email_id)
