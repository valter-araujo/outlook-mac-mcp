from datetime import UTC, datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.time_window import TimeWindow

LISBON = ZoneInfo("Europe/Lisbon")
MIDNIGHT = datetime(2026, 9, 15, tzinfo=LISBON)
NEXT_MIDNIGHT = MIDNIGHT + timedelta(days=1)
TODAY = TimeWindow(start=MIDNIGHT, end=NEXT_MIDNIGHT)


def test_rejects_a_window_that_ends_before_it_starts() -> None:
    with pytest.raises(InvalidRequestError):
        TimeWindow(start=NEXT_MIDNIGHT, end=MIDNIGHT)


def test_rejects_an_empty_window() -> None:
    with pytest.raises(InvalidRequestError):
        TimeWindow(start=MIDNIGHT, end=MIDNIGHT)


@pytest.mark.parametrize(
    ("start", "end"),
    [
        (datetime(2026, 9, 15), NEXT_MIDNIGHT),
        (MIDNIGHT, datetime(2026, 9, 16)),
    ],
)
def test_rejects_a_naive_boundary(start: datetime, end: datetime) -> None:
    with pytest.raises(InvalidRequestError):
        TimeWindow(start=start, end=end)


def test_compares_boundaries_as_instants_not_as_wall_clocks() -> None:
    """02:00 in Lisbon (UTC+1 in September) is 01:00 UTC, so it is not after 01:30 UTC."""
    with pytest.raises(InvalidRequestError):
        TimeWindow(
            start=datetime(2026, 9, 15, 1, 30, tzinfo=UTC),
            end=datetime(2026, 9, 15, 2, tzinfo=LISBON),
        )


def test_an_event_spanning_midnight_overlaps_both_days() -> None:
    start = MIDNIGHT - timedelta(hours=1)
    end = MIDNIGHT + timedelta(hours=1)
    yesterday = TimeWindow(start=MIDNIGHT - timedelta(days=1), end=MIDNIGHT)

    assert TODAY.overlaps(start, end)
    assert yesterday.overlaps(start, end)


def test_an_event_ending_exactly_at_the_start_is_outside() -> None:
    assert not TODAY.overlaps(MIDNIGHT - timedelta(hours=1), MIDNIGHT)


def test_an_all_day_event_belongs_to_one_day_only() -> None:
    tomorrow = TimeWindow(start=NEXT_MIDNIGHT, end=NEXT_MIDNIGHT + timedelta(days=1))

    assert TODAY.overlaps(MIDNIGHT, NEXT_MIDNIGHT)
    assert not tomorrow.overlaps(MIDNIGHT, NEXT_MIDNIGHT)


def test_an_event_inside_the_window_overlaps_it() -> None:
    assert TODAY.overlaps(MIDNIGHT + timedelta(hours=9), MIDNIGHT + timedelta(hours=10))
