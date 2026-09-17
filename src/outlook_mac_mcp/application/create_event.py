from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_draft import EventDraft
from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.domain.event import Event


class CreateEvent:
    """Second half of a write: the only step that changes the calendar.

    The draft is taken before the write, not after. If the write fails the user previews
    again, which costs a call; taking it afterwards would let a retry after a failure
    that happened past the wire create the same event twice.
    """

    def __init__(
        self, draft_store: DraftStore[EventDraft], calendar_writer: CalendarWriter
    ) -> None:
        self._draft_store = draft_store
        self._calendar_writer = calendar_writer

    def execute(self, token: str) -> Event:
        draft = self._draft_store.take(token)
        return self._calendar_writer.create(draft.new_event)
