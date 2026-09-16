from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.new_event import (
    MAX_ATTENDEES,
    MAX_BODY_LENGTH,
    MAX_SUBJECT_LENGTH,
    MIN_SUBJECT_LENGTH,
    NewEvent,
)

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
TEN = NINE + timedelta(hours=1)
MIDNIGHT = datetime(2026, 9, 15, tzinfo=SAO_PAULO)


def attendees(count: int) -> tuple[EmailAddress, ...]:
    return tuple(EmailAddress(address=f"person{index}@example.com") for index in range(count))


def test_accepts_a_well_formed_event() -> None:
    event = NewEvent(subject="Planning", start=NINE, end=TEN, location="Room 1")

    assert event.subject == "Planning"
    assert event.is_all_day is False
    assert event.attendees == ()


@pytest.mark.parametrize("length", [MIN_SUBJECT_LENGTH, MAX_SUBJECT_LENGTH])
def test_accepts_a_subject_at_the_edge_of_the_allowed_length(length: int) -> None:
    assert NewEvent(subject="a" * length, start=NINE, end=TEN).subject


@pytest.mark.parametrize("length", [MIN_SUBJECT_LENGTH - 1, MAX_SUBJECT_LENGTH + 1])
def test_rejects_a_subject_outside_the_allowed_length(length: int) -> None:
    with pytest.raises(InvalidRequestError):
        NewEvent(subject="a" * length, start=NINE, end=TEN)


def test_rejects_an_end_that_is_not_after_the_start() -> None:
    with pytest.raises(InvalidRequestError):
        NewEvent(subject="Planning", start=TEN, end=NINE)
    with pytest.raises(InvalidRequestError):
        NewEvent(subject="Planning", start=NINE, end=NINE)


def test_rejects_a_naive_boundary() -> None:
    with pytest.raises(InvalidRequestError):
        NewEvent(subject="Planning", start=datetime(2026, 9, 15, 9), end=TEN)


def test_accepts_the_maximum_number_of_attendees() -> None:
    event = NewEvent(subject="Planning", start=NINE, end=TEN, attendees=attendees(MAX_ATTENDEES))

    assert len(event.attendees) == MAX_ATTENDEES


def test_rejects_more_attendees_than_allowed() -> None:
    with pytest.raises(InvalidRequestError):
        NewEvent(subject="Planning", start=NINE, end=TEN, attendees=attendees(MAX_ATTENDEES + 1))


def test_accepts_an_all_day_event_running_midnight_to_midnight() -> None:
    event = NewEvent(
        subject="Offsite", start=MIDNIGHT, end=MIDNIGHT + timedelta(days=2), is_all_day=True
    )

    assert event.is_all_day is True


def test_rejects_an_all_day_event_with_a_clock_time() -> None:
    with pytest.raises(InvalidRequestError):
        NewEvent(subject="Offsite", start=NINE, end=MIDNIGHT + timedelta(days=1), is_all_day=True)


def test_defaults_to_no_body() -> None:
    assert NewEvent(subject="Planning", start=NINE, end=TEN).body == ""


def test_accepts_a_body_at_the_maximum_length() -> None:
    event = NewEvent(subject="Planning", start=NINE, end=TEN, body="a" * MAX_BODY_LENGTH)

    assert len(event.body) == MAX_BODY_LENGTH


def test_rejects_a_body_beyond_the_maximum_length() -> None:
    with pytest.raises(InvalidRequestError):
        NewEvent(subject="Planning", start=NINE, end=TEN, body="a" * (MAX_BODY_LENGTH + 1))
