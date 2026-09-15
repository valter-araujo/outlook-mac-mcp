from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.list_upcoming_events import (
    DEFAULT_UPCOMING_DAYS,
    MAX_UPCOMING_DAYS,
    MIN_UPCOMING_DAYS,
    ListUpcomingEvents,
    ListUpcomingEventsRequest,
)
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.event import Event
from tests.fakes.fixed_clock import FixedClock
from tests.fakes.in_memory_calendar_repository import InMemoryCalendarRepository

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NOW = datetime(2026, 9, 15, 15, 42, tzinfo=SAO_PAULO)


def make_event(event_id: str, start: datetime, end: datetime) -> Event:
    return Event(
        id=event_id,
        subject="Planning",
        start=start,
        end=end,
        is_all_day=False,
        location="",
        organizer=EmailAddress(address="ana@example.com"),
    )


def use_case_with(*events: Event) -> ListUpcomingEvents:
    repository = InMemoryCalendarRepository()
    for event in events:
        repository.add(event)
    return ListUpcomingEvents(repository, FixedClock(NOW))


def test_defaults_to_a_week() -> None:
    assert ListUpcomingEventsRequest().days == DEFAULT_UPCOMING_DAYS


@pytest.mark.parametrize("days", [MIN_UPCOMING_DAYS - 1, MAX_UPCOMING_DAYS + 1])
def test_rejects_days_outside_the_allowed_range(days: int) -> None:
    with pytest.raises(InvalidRequestError):
        ListUpcomingEventsRequest(days=days)


def test_returns_empty_when_nothing_is_coming_up() -> None:
    use_case = use_case_with(make_event("past", NOW - timedelta(days=1), NOW - timedelta(hours=23)))

    page = use_case.execute(ListUpcomingEventsRequest())

    assert page.items == ()
    assert page.total == 0
    assert page.total_is_exact is True


def test_returns_the_events_inside_the_requested_days_earliest_first() -> None:
    use_case = use_case_with(
        make_event("day-two", NOW + timedelta(days=2), NOW + timedelta(days=2, hours=1)),
        make_event("tomorrow", NOW + timedelta(days=1), NOW + timedelta(days=1, hours=1)),
        make_event("day-four", NOW + timedelta(days=4), NOW + timedelta(days=4, hours=1)),
    )

    page = use_case.execute(ListUpcomingEventsRequest(days=3))

    assert [event.id for event in page.items] == ["tomorrow", "day-two"]
    assert page.total == 2


def test_includes_an_event_already_in_progress() -> None:
    use_case = use_case_with(
        make_event("in-progress", NOW - timedelta(minutes=10), NOW + timedelta(minutes=50))
    )

    assert [event.id for event in use_case.execute(ListUpcomingEventsRequest()).items] == [
        "in-progress"
    ]


def test_excludes_an_event_that_has_already_ended_today() -> None:
    use_case = use_case_with(make_event("over", NOW - timedelta(hours=2), NOW - timedelta(hours=1)))

    assert use_case.execute(ListUpcomingEventsRequest()).items == ()


def test_the_window_closes_exactly_days_ahead() -> None:
    use_case = use_case_with(
        make_event("at-the-edge", NOW + timedelta(days=7), NOW + timedelta(days=7, hours=1))
    )

    assert use_case.execute(ListUpcomingEventsRequest(days=7)).items == ()
