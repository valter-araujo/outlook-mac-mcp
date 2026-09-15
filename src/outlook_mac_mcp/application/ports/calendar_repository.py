from typing import Protocol

from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.time_window import TimeWindow


class CalendarRepository(Protocol):
    def list_events(self, window: TimeWindow) -> Page[Event]:
        """Return every event overlapping `window`, earliest start first.

        Recurring series contribute their occurrences inside the window, not the series.
        The page's total may be a lower bound when the backend stops fetching at a cap.
        """
        ...
