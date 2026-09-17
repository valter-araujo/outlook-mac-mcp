from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_deletion_draft import EventDeletionDraft
from outlook_mac_mcp.application.preview_event_deletion import (
    EventDeletionRequest,
    PreviewEventDeletion,
)
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import (
    DraftNotFoundError,
    EventNotFoundError,
    InvalidRequestError,
)
from outlook_mac_mcp.domain.event import Event
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
    body="Bring the deck.",
    attendees=(EmailAddress(address="ana@example.com"),),
)


def new_store() -> DraftStore[EventDeletionDraft]:
    return DraftStore()


def test_parks_the_event_id_in_the_store_under_the_returned_token() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)

    draft = PreviewEventDeletion(writer, new_store()).execute(
        EventDeletionRequest(event_id=CURRENT.id)
    )

    assert draft.event_id == CURRENT.id


def test_the_summary_shows_the_full_current_event() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)

    draft = PreviewEventDeletion(writer, new_store()).execute(
        EventDeletionRequest(event_id=CURRENT.id)
    )

    assert 'subject: "Planning"' in draft.summary
    assert 'location: "Room 1"' in draft.summary
    assert 'body: "Bring the deck."' in draft.summary
    assert "attendees: ana@example.com" in draft.summary


def test_raises_when_the_event_does_not_exist() -> None:
    writer = InMemoryCalendarWriter()

    with pytest.raises(EventNotFoundError):
        PreviewEventDeletion(writer, new_store()).execute(
            EventDeletionRequest(event_id="never-seeded")
        )


def test_request_rejects_an_empty_event_id() -> None:
    with pytest.raises(InvalidRequestError):
        EventDeletionRequest(event_id="")


def test_a_token_from_here_is_not_found_in_a_different_stores_namespace() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    other_store: DraftStore[EventDeletionDraft] = DraftStore()

    draft = PreviewEventDeletion(writer, new_store()).execute(
        EventDeletionRequest(event_id=CURRENT.id)
    )

    with pytest.raises(DraftNotFoundError):
        other_store.take(draft.token)
