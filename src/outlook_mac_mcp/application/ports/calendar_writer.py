from typing import Protocol

from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.new_event import NewEvent


class CalendarWriter(Protocol):
    def create(self, new_event: NewEvent) -> Event:
        """Create `new_event` on the default calendar and return it as stored."""
        ...
