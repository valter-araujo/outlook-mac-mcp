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

    The deletion pair is optional independently of the other two: it is gated by its
    own flag on top of calendar write (see `Settings.calendar_delete_enabled`), so
    turning write on never silently enables the one irreversible action.
    """

    preview_event: PreviewEvent
    create_event: CreateEvent
    preview_event_update: PreviewEventUpdate
    update_event: UpdateEvent
    preview_event_deletion: PreviewEventDeletion | None
    delete_event: DeleteEvent | None
