from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.time_window import TimeWindow


class InMemoryCalendarRepository:
    """Fake CalendarRepository backed by a list; applies the port's overlap and order rules.

    Everything is in memory, so every total is exact.
    """

    def __init__(self) -> None:
        self._events: list[Event] = []

    def add(self, event: Event) -> None:
        self._events.append(event)

    def list_events(self, window: TimeWindow) -> Page[Event]:
        inside = [event for event in self._events if window.overlaps(event.start, event.end)]
        by_start = tuple(sorted(inside, key=lambda event: event.start))
        return Page(items=by_start, total=len(by_start), total_is_exact=True)
