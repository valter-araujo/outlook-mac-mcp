from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_change_summary import describe_changes, merge_changes
from outlook_mac_mcp.application.event_update_draft import EventUpdateDraft
from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.domain.event_changes import EventChanges


class PreviewEventUpdate:
    """First half of an update: nothing leaves the process here except the read that
    fetches the current event to diff against and to validate the merge with.

    Uses its own draft store, with its own token namespace: a token from here can never
    be handed to DeleteEvent or CreateEvent, and a create or deletion token is never
    accepted here.
    """

    def __init__(
        self, calendar_writer: CalendarWriter, draft_store: DraftStore[EventUpdateDraft]
    ) -> None:
        self._calendar_writer = calendar_writer
        self._draft_store = draft_store

    def execute(self, changes: EventChanges) -> EventUpdateDraft:
        current = self._calendar_writer.get_by_id(changes.event_id)
        merge_changes(current, changes)
        summary = describe_changes(current, changes)
        return self._draft_store.add(
            lambda token: EventUpdateDraft(token=token, summary=summary, changes=changes)
        )
