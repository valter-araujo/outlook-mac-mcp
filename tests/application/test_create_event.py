from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.create_event import CreateEvent
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_draft import EventDraft
from outlook_mac_mcp.domain.errors import DraftNotFoundError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.new_event import NewEvent
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
from tests.fakes.in_memory_calendar_writer import InMemoryCalendarWriter

NINE = datetime(2026, 9, 15, 9, tzinfo=ZoneInfo("America/Sao_Paulo"))
AN_EVENT = NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1))


class FailingCalendarWriter:
    def __init__(self) -> None:
        self.calls = 0

    def create(self, new_event: NewEvent) -> Event:
        self.calls += 1
        raise GraphRequestError("Graph returned 503", 503, "ServiceUnavailable")


def new_store() -> DraftStore[EventDraft]:
    return DraftStore()


def add_draft(store: DraftStore[EventDraft], new_event: NewEvent, summary: str) -> EventDraft:
    return store.add(lambda token: EventDraft(token=token, summary=summary, new_event=new_event))


def test_creates_the_drafted_event_and_returns_it_as_stored() -> None:
    store = new_store()
    writer = InMemoryCalendarWriter()
    token = add_draft(store, AN_EVENT, "summary").token

    created = CreateEvent(store, writer).execute(token)

    assert writer.created == [AN_EVENT]
    assert created.id == "created-1"
    assert created.subject == "Planning"


def test_a_token_creates_at_most_one_event() -> None:
    store = new_store()
    writer = InMemoryCalendarWriter()
    token = add_draft(store, AN_EVENT, "summary").token
    use_case = CreateEvent(store, writer)
    use_case.execute(token)

    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert len(writer.created) == 1


def test_an_unknown_token_creates_nothing() -> None:
    writer = InMemoryCalendarWriter()

    with pytest.raises(DraftNotFoundError):
        CreateEvent(new_store(), writer).execute("never-issued")

    assert writer.created == []


def test_a_failed_write_consumes_the_draft_so_a_retry_cannot_duplicate() -> None:
    store = new_store()
    writer = FailingCalendarWriter()
    token = add_draft(store, AN_EVENT, "summary").token
    use_case = CreateEvent(store, writer)

    with pytest.raises(GraphRequestError):
        use_case.execute(token)
    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert writer.calls == 1
