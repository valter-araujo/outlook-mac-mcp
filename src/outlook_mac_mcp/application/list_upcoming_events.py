from dataclasses import dataclass
from datetime import timedelta

from outlook_mac_mcp.application.clock import Clock
from outlook_mac_mcp.application.ports.calendar_repository import CalendarRepository
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.time_window import TimeWindow

MIN_UPCOMING_DAYS = 1
MAX_UPCOMING_DAYS = 30
DEFAULT_UPCOMING_DAYS = 7


@dataclass(frozen=True, slots=True)
class ListUpcomingEventsRequest:
    days: int = DEFAULT_UPCOMING_DAYS

    def __post_init__(self) -> None:
        if not MIN_UPCOMING_DAYS <= self.days <= MAX_UPCOMING_DAYS:
            raise InvalidRequestError(
                f"days must be between {MIN_UPCOMING_DAYS} and {MAX_UPCOMING_DAYS}"
            )


class ListUpcomingEvents:
    """Upcoming starts now, not at midnight: a meeting that began ten minutes ago is still
    in progress and still listed, while this morning's finished ones are not.
    """

    def __init__(self, calendar_repository: CalendarRepository, clock: Clock) -> None:
        self._calendar_repository = calendar_repository
        self._clock = clock

    def execute(self, request: ListUpcomingEventsRequest) -> Page[Event]:
        now = self._clock.now()
        window = TimeWindow(start=now, end=now + timedelta(days=request.days))
        return self._calendar_repository.list_events(window)
