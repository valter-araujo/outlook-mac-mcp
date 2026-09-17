from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.event_changes import MAX_EVENT_ID_LENGTH, EventChanges

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
TEN = NINE + timedelta(hours=1)


def test_accepts_a_single_supplied_field() -> None:
    changes = EventChanges(event_id="AAMkNEW", subject="Rescheduled")

    assert changes.subject == "Rescheduled"
    assert changes.start is None
    assert changes.has_any_change() is True


def test_rejects_an_empty_event_id() -> None:
    with pytest.raises(InvalidRequestError):
        EventChanges(event_id="", subject="Rescheduled")


def test_rejects_an_event_id_beyond_the_maximum_length() -> None:
    with pytest.raises(InvalidRequestError):
        EventChanges(event_id="a" * (MAX_EVENT_ID_LENGTH + 1), subject="Rescheduled")


def test_accepts_an_event_id_at_the_maximum_length() -> None:
    changes = EventChanges(event_id="a" * MAX_EVENT_ID_LENGTH, subject="Rescheduled")

    assert changes.event_id


def test_rejects_a_naive_start() -> None:
    with pytest.raises(InvalidRequestError):
        EventChanges(event_id="AAMkNEW", start=datetime(2026, 9, 15, 9))


def test_rejects_a_naive_end() -> None:
    with pytest.raises(InvalidRequestError):
        EventChanges(event_id="AAMkNEW", end=datetime(2026, 9, 15, 10))


def test_accepts_start_alone_without_needing_the_current_end() -> None:
    """Whether this leaves the merged event valid depends on the current end, which
    EventChanges does not know; that check happens at the merge step, not here.
    """
    changes = EventChanges(event_id="AAMkNEW", start=NINE)

    assert changes.start == NINE


def test_rejects_an_end_before_the_start_when_both_are_supplied() -> None:
    with pytest.raises(InvalidRequestError):
        EventChanges(event_id="AAMkNEW", start=TEN, end=NINE)


def test_accepts_start_before_end_when_both_are_supplied() -> None:
    changes = EventChanges(event_id="AAMkNEW", start=NINE, end=TEN)

    assert changes.start == NINE
    assert changes.end == TEN


def test_rejects_no_fields_supplied_at_all() -> None:
    with pytest.raises(InvalidRequestError):
        EventChanges(event_id="AAMkNEW")
