from dataclasses import dataclass

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_deletion_draft import EventDeletionDraft
from outlook_mac_mcp.application.event_deletion_summary import describe_event_for_deletion
from outlook_mac_mcp.application.ports.calendar_writer import CalendarWriter
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.event_changes import MAX_EVENT_ID_LENGTH


@dataclass(frozen=True, slots=True)
class EventDeletionRequest:
    event_id: str

    def __post_init__(self) -> None:
        if not self.event_id:
            raise InvalidRequestError("event_id must not be empty")
        if len(self.event_id) > MAX_EVENT_ID_LENGTH:
            raise InvalidRequestError(f"event_id must be at most {MAX_EVENT_ID_LENGTH} characters")


class PreviewEventDeletion:
    """First half of a deletion: nothing leaves the process here except the read that
    fetches the event the confirmation will show in full.

    Uses its own draft store, with its own token namespace: a token from here can never
    be handed to UpdateEvent or CreateEvent, and an update or create token is never
    accepted here.
    """

    def __init__(
        self, calendar_writer: CalendarWriter, draft_store: DraftStore[EventDeletionDraft]
    ) -> None:
        self._calendar_writer = calendar_writer
        self._draft_store = draft_store

    def execute(self, request: EventDeletionRequest) -> EventDeletionDraft:
        event = self._calendar_writer.get_by_id(request.event_id)
        summary = describe_event_for_deletion(event)
        return self._draft_store.add(
            lambda token: EventDeletionDraft(token=token, summary=summary, event_id=event.id)
        )
