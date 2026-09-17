from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_deletion_draft import EventDeletionDraft
from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter


class DeleteEvent:
    """Second half of a deletion: the only step that changes the calendar, and the only
    one that cannot be undone by previewing again.

    The draft is taken before the write, not after, the same ordering CreateEvent and
    UpdateEvent use and for the same reason: a failed write costs a re-preview instead of
    risking a retry that could delete an event a concurrent change already replaced.
    """

    def __init__(
        self, draft_store: DraftStore[EventDeletionDraft], calendar_writer: CalendarWriter
    ) -> None:
        self._draft_store = draft_store
        self._calendar_writer = calendar_writer

    def execute(self, token: str) -> str:
        draft = self._draft_store.take(token)
        self._calendar_writer.delete(draft.event_id)
        return draft.event_id
