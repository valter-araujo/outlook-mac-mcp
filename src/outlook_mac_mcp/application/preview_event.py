from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_draft import EventDraft
from outlook_mac_mcp.application.event_summary import describe
from outlook_mac_mcp.domain.new_event import NewEvent


class PreviewEvent:
    """First half of a write: nothing leaves the process here.

    The event is already validated by construction; this step only parks it under a
    token and describes it, so the user confirms exactly what CreateEvent will send.
    """

    def __init__(self, draft_store: DraftStore[EventDraft]) -> None:
        self._draft_store = draft_store

    def execute(self, new_event: NewEvent) -> EventDraft:
        return self._draft_store.add(
            lambda token: EventDraft(token=token, summary=describe(new_event), new_event=new_event)
        )
