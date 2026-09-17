from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_update_draft import EventUpdateDraft
from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.domain.event import Event


class UpdateEvent:
    """Second half of an update: the only step that changes the calendar.

    The draft is taken before the write, not after, the same ordering CreateEvent uses
    and for the same reason: a failed write costs a re-preview instead of risking a
    retry applying the same patch twice from stale state.
    """

    def __init__(
        self, draft_store: DraftStore[EventUpdateDraft], calendar_writer: CalendarWriter
    ) -> None:
        self._draft_store = draft_store
        self._calendar_writer = calendar_writer

    def execute(self, token: str) -> Event:
        draft = self._draft_store.take(token)
        return self._calendar_writer.update(draft.changes)
