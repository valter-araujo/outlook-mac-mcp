from datetime import timedelta

from outlook_mac_mcp.application.clock import Clock
from outlook_mac_mcp.application.ports.calendar_repository import CalendarRepository
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.time_window import TimeWindow

ONE_DAY = timedelta(days=1)


class ListTodaysEvents:
    """Today is midnight to the next midnight in the clock's zone.

    Adding a day to an aware datetime is wall-clock arithmetic in Python, so the window
    ends at tomorrow's midnight even when a DST change makes the day 23 or 25 hours long.
    """

    def __init__(self, calendar_repository: CalendarRepository, clock: Clock) -> None:
        self._calendar_repository = calendar_repository
        self._clock = clock

    def execute(self) -> Page[Event]:
        start_of_today = self._clock.now().replace(hour=0, minute=0, second=0, microsecond=0)
        window = TimeWindow(start=start_of_today, end=start_of_today + ONE_DAY)
        return self._calendar_repository.list_events(window)
