from dataclasses import dataclass

from outlook_mac_mcp.application.create_event import CreateEvent
from outlook_mac_mcp.application.preview_event import PreviewEvent


@dataclass(frozen=True, slots=True)
class CalendarWriteUseCases:
    """The write pair, bundled so the server can register both or neither."""

    preview_event: PreviewEvent
    create_event: CreateEvent
