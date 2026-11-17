from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class DraftDeletionPreview:
    """A pending draft deletion waiting for the user's go-ahead."""

    token: str
    summary: str
    draft_id: str
