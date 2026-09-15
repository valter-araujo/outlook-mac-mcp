from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from outlook_mac_mcp.application.list_todays_events import ListTodaysEvents
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event import Event
from tests.fakes.fixed_clock import FixedClock
from tests.fakes.in_memory_calendar_repository import InMemoryCalendarRepository

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
MIDNIGHT = datetime(2026, 9, 15, tzinfo=SAO_PAULO)
MID_AFTERNOON = MIDNIGHT + timedelta(hours=15, minutes=42)


def make_event(event_id: str, start: datetime, end: datetime, *, is_all_day: bool = False) -> Event:
    return Event(
        id=event_id,
        subject="Planning",
        start=start,
        end=end,
        is_all_day=is_all_day,
        location="Room 1",
        organizer=EmailAddress(address="ana@example.com", display_name="Ana Lima"),
    )


def hours(count: float) -> timedelta:
    return timedelta(hours=count)


def use_case_with(*events: Event, now: datetime = MID_AFTERNOON) -> ListTodaysEvents:
    repository = InMemoryCalendarRepository()
    for event in events:
        repository.add(event)
    return ListTodaysEvents(repository, FixedClock(now))


def test_returns_empty_when_nothing_is_scheduled_today() -> None:
    use_case = use_case_with(make_event("tomorrow", MIDNIGHT + hours(33), MIDNIGHT + hours(34)))

    page = use_case.execute()

    assert page.items == ()
    assert page.total == 0
    assert page.total_is_exact is True


def test_returns_todays_events_earliest_first() -> None:
    use_case = use_case_with(
        make_event("later", MIDNIGHT + hours(14), MIDNIGHT + hours(15)),
        make_event("earlier", MIDNIGHT + hours(9), MIDNIGHT + hours(10)),
        make_event("yesterday", MIDNIGHT - hours(5), MIDNIGHT - hours(4)),
    )

    page = use_case.execute()

    assert [event.id for event in page.items] == ["earlier", "later"]
    assert page.total == 2


def test_includes_an_event_that_started_yesterday_and_ends_today() -> None:
    use_case = use_case_with(make_event("overnight", MIDNIGHT - hours(1), MIDNIGHT + hours(1)))

    assert [event.id for event in use_case.execute().items] == ["overnight"]


def test_includes_an_event_that_starts_today_and_ends_tomorrow() -> None:
    use_case = use_case_with(make_event("late", MIDNIGHT + hours(23), MIDNIGHT + hours(25)))

    assert [event.id for event in use_case.execute().items] == ["late"]


def test_includes_todays_all_day_event_but_not_tomorrows() -> None:
    use_case = use_case_with(
        make_event("today", MIDNIGHT, MIDNIGHT + hours(24), is_all_day=True),
        make_event("tomorrow", MIDNIGHT + hours(24), MIDNIGHT + hours(48), is_all_day=True),
    )

    assert [event.id for event in use_case.execute().items] == ["today"]


def test_today_is_decided_by_the_clock_not_by_the_events() -> None:
    use_case = use_case_with(
        make_event("today", MIDNIGHT + hours(9), MIDNIGHT + hours(10)),
        now=MIDNIGHT + hours(24) + timedelta(minutes=1),
    )

    assert use_case.execute().items == ()


def test_today_runs_until_midnight_even_when_asked_late_in_the_day() -> None:
    use_case = use_case_with(
        make_event("last-slot", MIDNIGHT + hours(23.5), MIDNIGHT + hours(23.75)),
        now=MIDNIGHT + hours(23),
    )

    assert [event.id for event in use_case.execute().items] == ["last-slot"]
