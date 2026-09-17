from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_update_draft import EventUpdateDraft
from outlook_mac_mcp.application.preview_event_update import PreviewEventUpdate
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import (
    DraftNotFoundError,
    EventNotFoundError,
    InvalidRequestError,
)
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from tests.fakes.in_memory_calendar_writer import InMemoryCalendarWriter

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
)


def new_store() -> DraftStore[EventUpdateDraft]:
    return DraftStore()


def test_parks_the_changes_in_the_store_under_the_returned_token() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    changes = EventChanges(event_id=CURRENT.id, subject="Replanning")

    draft = PreviewEventUpdate(writer, new_store()).execute(changes)

    assert draft.changes == changes


def test_the_summary_shows_the_diff_not_just_the_resulting_state() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    changes = EventChanges(event_id=CURRENT.id, subject="Replanning")

    draft = PreviewEventUpdate(writer, new_store()).execute(changes)

    assert 'subject: "Planning" -> "Replanning"' in draft.summary


def test_raises_when_the_event_does_not_exist() -> None:
    writer = InMemoryCalendarWriter()
    changes = EventChanges(event_id="never-seeded", subject="Replanning")

    with pytest.raises(EventNotFoundError):
        PreviewEventUpdate(writer, new_store()).execute(changes)


def test_raises_when_the_merged_event_would_violate_a_new_event_rule() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    changes = EventChanges(event_id=CURRENT.id, subject="a" * 1000)

    with pytest.raises(InvalidRequestError):
        PreviewEventUpdate(writer, new_store()).execute(changes)


def test_a_token_from_here_is_not_found_in_a_different_stores_namespace() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    other_store: DraftStore[EventUpdateDraft] = DraftStore()
    changes = EventChanges(event_id=CURRENT.id, subject="Replanning")

    draft = PreviewEventUpdate(writer, new_store()).execute(changes)

    with pytest.raises(DraftNotFoundError):
        other_store.take(draft.token)
