from dataclasses import dataclass

from outlook_mac_mcp.application.draft_deletion_preview import DraftDeletionPreview
from outlook_mac_mcp.application.draft_deletion_summary import describe_draft_for_deletion
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.get_email import MAX_EMAIL_ID_LENGTH
from outlook_mac_mcp.application.ports.draft_deleter import DraftDeleter
from outlook_mac_mcp.domain.errors import InvalidRequestError


@dataclass(frozen=True, slots=True)
class DraftDeletionRequest:
    email_id: str

    def __post_init__(self) -> None:
        if not self.email_id:
            raise InvalidRequestError("email_id must not be empty")
        if len(self.email_id) > MAX_EMAIL_ID_LENGTH:
            raise InvalidRequestError(f"email_id must be at most {MAX_EMAIL_ID_LENGTH} characters")


class PreviewDraftDeletion:
    """First half of a draft deletion: confirms the message is a draft in Drafts, writes
    nothing.

    Refusing here, rather than only at delete time, means a token is never issued for a
    message the deletion would refuse anyway.
    """

    def __init__(
        self, draft_deleter: DraftDeleter, draft_store: DraftStore[DraftDeletionPreview]
    ) -> None:
        self._draft_deleter = draft_deleter
        self._draft_store = draft_store

    def execute(self, request: DraftDeletionRequest) -> DraftDeletionPreview:
        draft = self._draft_deleter.get_draft(request.email_id)
        summary = describe_draft_for_deletion(draft)
        return self._draft_store.add(
            lambda token: DraftDeletionPreview(token=token, summary=summary, draft_id=draft.id)
        )
