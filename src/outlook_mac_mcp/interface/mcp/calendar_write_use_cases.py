from dataclasses import dataclass

from outlook_mac_mcp.application.create_event import CreateEvent
from outlook_mac_mcp.application.delete_event import DeleteEvent
from outlook_mac_mcp.application.preview_event import PreviewEvent
from outlook_mac_mcp.application.preview_event_deletion import PreviewEventDeletion
from outlook_mac_mcp.application.preview_event_update import PreviewEventUpdate
from outlook_mac_mcp.application.update_event import UpdateEvent


@dataclass(frozen=True, slots=True)
class CalendarWriteUseCases:
    """The three preview/apply pairs, bundled so the server can register all of them or
    none. Each pair carries its own draft store, so a token from one can never be
    accepted by another.
    """

    preview_event: PreviewEvent
    create_event: CreateEvent
    preview_event_update: PreviewEventUpdate
    update_event: UpdateEvent
    preview_event_deletion: PreviewEventDeletion
    delete_event: DeleteEvent
