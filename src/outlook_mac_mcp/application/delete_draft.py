from outlook_mac_mcp.application.draft_deletion_preview import DraftDeletionPreview
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.ports.draft_deleter import DraftDeleter


class DeleteDraft:
    """Second half of a draft deletion: the only step that changes the mailbox.

    The draft is checked again right before the delete, unlike the calendar deletion,
    because "only drafts in Drafts" is the whole safety rule of this tool: if the user
    sent the draft or moved it since the preview, the token must not delete a message
    that is no longer a draft.
    """

    def __init__(
        self, draft_store: DraftStore[DraftDeletionPreview], draft_deleter: DraftDeleter
    ) -> None:
        self._draft_store = draft_store
        self._draft_deleter = draft_deleter

    def execute(self, token: str) -> str:
        preview = self._draft_store.take(token)
        draft = self._draft_deleter.get_draft(preview.draft_id)
        self._draft_deleter.delete(draft.id)
        return draft.id
