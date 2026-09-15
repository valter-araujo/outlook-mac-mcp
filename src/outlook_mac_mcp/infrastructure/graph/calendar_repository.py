from collections.abc import Mapping
from typing import Any
from zoneinfo import ZoneInfo

from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.time_window import TimeWindow
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.event_mapper import EVENT_FIELDS, to_event
from outlook_mac_mcp.infrastructure.graph.pagination import read_items, read_next_link
from outlook_mac_mcp.infrastructure.graph.preferences import timezone_preference

CALENDAR_VIEW_PATH = "/me/calendarView"
START_FIRST_ORDER = "start/dateTime"
PAGE_SIZE = 100
# Ten pages of a hundred is far beyond a month of any real calendar, so the cap exists
# to bound a runaway series, not to be reached; past it the total is a lower bound.
MAX_PAGES = 10


class GraphCalendarRepository:
    """Reads the calendar view from Microsoft Graph for the signed-in account.

    calendarView, not events: the view expands recurring series into the occurrences
    inside the window, which is what a day or a week actually holds.

    Satisfies the CalendarRepository port structurally.
    """

    def __init__(self, client: GraphClient, timezone: ZoneInfo) -> None:
        self._client = client
        self._timezone = timezone

    def list_events(self, window: TimeWindow) -> Page[Event]:
        """The window goes out as ISO 8601 with offsets, which Graph reads as given, and the
        Prefer header brings every time back in the resolved zone so the mapper never
        has to convert.
        """
        payload = self._client.get(
            CALENDAR_VIEW_PATH,
            {
                "startDateTime": window.start.isoformat(),
                "endDateTime": window.end.isoformat(),
                "$select": ",".join(EVENT_FIELDS),
                "$orderby": START_FIRST_ORDER,
                "$top": PAGE_SIZE,
            },
            timezone_preference(self._timezone),
        )
        return self._read_all_pages(payload)

    def _read_all_pages(self, first_page: Mapping[str, Any]) -> Page[Event]:
        events = [to_event(item) for item in read_items(first_page)]
        next_link = read_next_link(first_page)
        pages_read = 1
        while next_link is not None and pages_read < MAX_PAGES:
            page = self._client.follow(next_link)
            events.extend(to_event(item) for item in read_items(page))
            next_link = read_next_link(page)
            pages_read += 1
        return Page(items=tuple(events), total=len(events), total_is_exact=next_link is None)
