import secrets

from outlook_mac_mcp.application.event_draft import EventDraft
from outlook_mac_mcp.domain.errors import DraftNotFoundError
from outlook_mac_mcp.domain.new_event import NewEvent

TOKEN_BYTES = 16


class DraftStore:
    """Drafts held in memory for the life of the server process.

    Not persisted on purpose: a draft is a pending confirmation, and a confirmation
    should not outlive the conversation that asked for it. A restart clears them all,
    and the client simply previews again.

    Tokens are unguessable rather than sequential so a create call cannot reach a draft
    it was not given.
    """

    def __init__(self) -> None:
        self._drafts: dict[str, EventDraft] = {}

    def add(self, new_event: NewEvent, summary: str) -> EventDraft:
        draft = EventDraft(
            token=secrets.token_urlsafe(TOKEN_BYTES), summary=summary, new_event=new_event
        )
        self._drafts[draft.token] = draft
        return draft

    def take(self, token: str) -> EventDraft:
        """Return the draft and forget it, so a token creates at most one event."""
        try:
            return self._drafts.pop(token)
        except KeyError:
            raise DraftNotFoundError("no pending draft for that token; preview again") from None
