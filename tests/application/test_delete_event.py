from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.delete_event import DeleteEvent
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_deletion_draft import EventDeletionDraft
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import DraftNotFoundError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
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


class FailingCalendarWriter:
    """Only `delete` is exercised here; the other CalendarWriter methods are unused by
    DeleteEvent but still need a body to satisfy the port.
    """

    def __init__(self) -> None:
        self.calls = 0

    def create(self, new_event: NewEvent) -> Event:
        raise NotImplementedError

    def get_by_id(self, event_id: str) -> Event:
        raise NotImplementedError

    def update(self, changes: EventChanges) -> Event:
        raise NotImplementedError

    def delete(self, event_id: str) -> None:
        self.calls += 1
        raise GraphRequestError("Graph returned 503", 503, "ServiceUnavailable")


def new_store() -> DraftStore[EventDeletionDraft]:
    return DraftStore()


def add_draft(store: DraftStore[EventDeletionDraft], event_id: str, summary: str) -> str:
    draft = store.add(
        lambda token: EventDeletionDraft(token=token, summary=summary, event_id=event_id)
    )
    return draft.token


def test_deletes_the_drafted_event() -> None:
    store = new_store()
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    token = add_draft(store, CURRENT.id, "summary")

    DeleteEvent(store, writer).execute(token)

    assert writer.deleted == [CURRENT.id]


def test_a_token_deletes_at_most_once() -> None:
    store = new_store()
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    token = add_draft(store, CURRENT.id, "summary")
    use_case = DeleteEvent(store, writer)
    use_case.execute(token)

    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert len(writer.deleted) == 1


def test_an_unknown_token_deletes_nothing() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)

    with pytest.raises(DraftNotFoundError):
        DeleteEvent(new_store(), writer).execute("never-issued")

    assert writer.deleted == []


def test_a_failed_delete_consumes_the_draft_so_a_retry_cannot_reapply_it() -> None:
    store = new_store()
    writer = FailingCalendarWriter()
    token = add_draft(store, CURRENT.id, "summary")
    use_case = DeleteEvent(store, writer)

    with pytest.raises(GraphRequestError):
        use_case.execute(token)
    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert writer.calls == 1
