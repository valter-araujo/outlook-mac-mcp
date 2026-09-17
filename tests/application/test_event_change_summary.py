from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.event_change_summary import describe_changes, merge_changes
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.sensitivity import Sensitivity
from outlook_mac_mcp.domain.show_as import ShowAs

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
ORGANIZER = EmailAddress(address="me@example.com", display_name="Me")
CURRENT = Event(
    id="AAMkEXISTING",
    subject="Planning",
    start=NINE,
    end=NINE + timedelta(hours=1),
    is_all_day=False,
    location="Room 1",
    organizer=ORGANIZER,
    body="Bring the deck.",
    attendees=(EmailAddress(address="ana@example.com"),),
)
CURRENT_WITH_EXTRAS = Event(
    id="AAMkEXISTING",
    subject="Planning",
    start=NINE,
    end=NINE + timedelta(hours=1),
    is_all_day=False,
    location="Room 1",
    organizer=ORGANIZER,
    reminder_minutes_before_start=15,
    sensitivity=Sensitivity.NORMAL,
    show_as=ShowAs.BUSY,
)


def test_merge_carries_over_every_field_left_unsupplied() -> None:
    changes = EventChanges(event_id=CURRENT.id, subject="Replanning")

    merged = merge_changes(CURRENT, changes)

    assert merged.subject == "Replanning"
    assert merged.start == CURRENT.start
    assert merged.end == CURRENT.end
    assert merged.location == CURRENT.location
    assert merged.body == CURRENT.body
    assert merged.attendees == CURRENT.attendees
    assert merged.reminder_minutes_before_start == CURRENT.reminder_minutes_before_start
    assert merged.sensitivity == CURRENT.sensitivity
    assert merged.show_as == CURRENT.show_as


def test_merge_carries_over_a_reminder_sensitivity_and_show_as_left_unsupplied() -> None:
    changes = EventChanges(event_id=CURRENT_WITH_EXTRAS.id, subject="Replanning")

    merged = merge_changes(CURRENT_WITH_EXTRAS, changes)

    assert merged.reminder_minutes_before_start == 15
    assert merged.sensitivity is Sensitivity.NORMAL
    assert merged.show_as is ShowAs.BUSY


def test_merge_applies_a_reminder_sensitivity_and_show_as_that_are_supplied() -> None:
    changes = EventChanges(
        event_id=CURRENT_WITH_EXTRAS.id,
        reminder_minutes_before_start=30,
        sensitivity=Sensitivity.PRIVATE,
        show_as=ShowAs.TENTATIVE,
    )

    merged = merge_changes(CURRENT_WITH_EXTRAS, changes)

    assert merged.reminder_minutes_before_start == 30
    assert merged.sensitivity is Sensitivity.PRIVATE
    assert merged.show_as is ShowAs.TENTATIVE


def test_merge_applies_every_field_that_is_supplied() -> None:
    new_start = NINE + timedelta(hours=2)
    changes = EventChanges(
        event_id=CURRENT.id,
        subject="Replanning",
        start=new_start,
        end=new_start + timedelta(hours=1),
        location="Room 2",
        body="Bring the new deck.",
        attendees=(EmailAddress(address="bo@example.com"),),
    )

    merged = merge_changes(CURRENT, changes)

    assert merged.subject == "Replanning"
    assert merged.start == new_start
    assert merged.location == "Room 2"
    assert merged.body == "Bring the new deck."
    assert merged.attendees == (EmailAddress(address="bo@example.com"),)


def test_merge_keeps_is_all_day_from_the_current_event_since_it_is_never_changeable() -> None:
    all_day_current = Event(
        id="AAMkALLDAY",
        subject="Offsite",
        start=datetime(2026, 9, 15, tzinfo=SAO_PAULO),
        end=datetime(2026, 9, 16, tzinfo=SAO_PAULO),
        is_all_day=True,
        location="",
        organizer=ORGANIZER,
    )
    changes = EventChanges(event_id=all_day_current.id, subject="Offsite retreat")

    merged = merge_changes(all_day_current, changes)

    assert merged.is_all_day is True


def test_merge_raises_when_the_resulting_event_violates_a_new_event_rule() -> None:
    changes = EventChanges(event_id=CURRENT.id, subject="a" * 1000)

    with pytest.raises(InvalidRequestError):
        merge_changes(CURRENT, changes)


def test_merge_raises_when_only_the_changed_boundary_breaks_start_before_end() -> None:
    changes = EventChanges(event_id=CURRENT.id, start=CURRENT.end + timedelta(hours=1))

    with pytest.raises(InvalidRequestError):
        merge_changes(CURRENT, changes)


def test_diff_shows_only_the_fields_that_were_supplied() -> None:
    changes = EventChanges(event_id=CURRENT.id, subject="Replanning")

    summary = describe_changes(CURRENT, changes)

    assert 'subject: "Planning" -> "Replanning"' in summary
    assert "start:" not in summary
    assert "location:" not in summary
    assert "body:" not in summary
    assert "attendees:" not in summary
    assert "reminder:" not in summary
    assert "sensitivity:" not in summary
    assert "show as:" not in summary


def test_diff_leads_with_the_current_subject_and_time_to_identify_the_event() -> None:
    changes = EventChanges(event_id=CURRENT.id, location="Room 2")

    summary = describe_changes(CURRENT, changes)

    assert summary.startswith('"Planning" (')


def test_diff_shows_a_line_even_when_the_new_value_equals_the_current_one() -> None:
    changes = EventChanges(event_id=CURRENT.id, subject="Planning")

    summary = describe_changes(CURRENT, changes)

    assert 'subject: "Planning" -> "Planning"' in summary


def test_diff_renders_empty_location_and_attendees_as_none() -> None:
    bare_current = Event(
        id="AAMkBARE",
        subject="Sync",
        start=NINE,
        end=NINE + timedelta(hours=1),
        is_all_day=False,
        location="",
        organizer=ORGANIZER,
    )
    changes = EventChanges(event_id=bare_current.id, location="Room 3", attendees=())

    summary = describe_changes(bare_current, changes)

    assert 'location: (none) -> "Room 3"' in summary
    assert "attendees: (none) -> (none)" in summary


def test_diff_shows_every_supplied_field_together() -> None:
    new_start = NINE + timedelta(hours=2)
    changes = EventChanges(
        event_id=CURRENT.id,
        subject="Replanning",
        start=new_start,
        end=new_start + timedelta(hours=1),
        location="Room 2",
        body="Bring the new deck.",
        attendees=(EmailAddress(address="bo@example.com"),),
    )

    summary = describe_changes(CURRENT, changes)

    assert 'subject: "Planning" -> "Replanning"' in summary
    assert "start:" in summary
    assert "end:" in summary
    assert 'location: "Room 1" -> "Room 2"' in summary
    assert 'body: "Bring the deck." -> "Bring the new deck."' in summary
    assert "attendees: ana@example.com -> bo@example.com" in summary


def test_diff_shows_a_changed_reminder_sensitivity_and_show_as() -> None:
    changes = EventChanges(
        event_id=CURRENT_WITH_EXTRAS.id,
        reminder_minutes_before_start=30,
        sensitivity=Sensitivity.PRIVATE,
        show_as=ShowAs.TENTATIVE,
    )

    summary = describe_changes(CURRENT_WITH_EXTRAS, changes)

    assert "reminder: 15 min before -> 30 min before" in summary
    assert "sensitivity: normal -> private" in summary
    assert "show as: busy -> tentative" in summary


def test_diff_renders_an_unset_reminder_sensitivity_and_show_as_as_none() -> None:
    changes = EventChanges(
        event_id=CURRENT.id,
        reminder_minutes_before_start=10,
        sensitivity=Sensitivity.CONFIDENTIAL,
        show_as=ShowAs.OOF,
    )

    summary = describe_changes(CURRENT, changes)

    assert "reminder: (none) -> 10 min before" in summary
    assert "sensitivity: (none) -> confidential" in summary
    assert "show as: (none) -> oof" in summary
