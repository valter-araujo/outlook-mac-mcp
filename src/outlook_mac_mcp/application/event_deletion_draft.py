from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class EventDeletionDraft:
    """A pending deletion waiting for the user's go-ahead.

    Only the id is kept for the delete call itself; the full event the summary
    describes was already fetched once, to build that summary, and is not re-verified
    at apply time, the same as CreateEvent and UpdateEvent never re-verify their drafts.
    """

    token: str
    summary: str
    event_id: str
