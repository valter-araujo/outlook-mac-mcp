from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_update_draft import EventUpdateDraft
from outlook_mac_mcp.application.update_event import UpdateEvent
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import DraftNotFoundError
from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.event_changes import EventChanges
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
A_CHANGE = EventChanges(event_id=CURRENT.id, subject="Replanning")


class FailingCalendarWriter:
    """Only `update` is exercised here; the other CalendarWriter methods are unused by
    UpdateEvent but still need a body to satisfy the port.
    """

    def __init__(self) -> None:
        self.calls = 0

    def create(self, new_event: object) -> Event:
        raise NotImplementedError

    def get_by_id(self, event_id: str) -> Event:
        raise NotImplementedError

    def update(self, changes: EventChanges) -> Event:
        self.calls += 1
        raise GraphRequestError("Graph returned 503", 503, "ServiceUnavailable")


def new_store() -> DraftStore[EventUpdateDraft]:
    return DraftStore()


def add_draft(store: DraftStore[EventUpdateDraft], changes: EventChanges, summary: str) -> str:
    draft = store.add(lambda token: EventUpdateDraft(token=token, summary=summary, changes=changes))
    return draft.token


def test_applies_the_drafted_changes_and_returns_the_event_as_stored() -> None:
    store = new_store()
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    token = add_draft(store, A_CHANGE, "summary")

    updated = UpdateEvent(store, writer).execute(token)

    assert writer.updated == [A_CHANGE]
    assert updated.subject == "Replanning"


def test_a_token_updates_at_most_once() -> None:
    store = new_store()
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)
    token = add_draft(store, A_CHANGE, "summary")
    use_case = UpdateEvent(store, writer)
    use_case.execute(token)

    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert len(writer.updated) == 1


def test_an_unknown_token_updates_nothing() -> None:
    writer = InMemoryCalendarWriter()
    writer.seed(CURRENT)

    with pytest.raises(DraftNotFoundError):
        UpdateEvent(new_store(), writer).execute("never-issued")

    assert writer.updated == []


def test_a_failed_write_consumes_the_draft_so_a_retry_cannot_reapply_it() -> None:
    store = new_store()
    writer = FailingCalendarWriter()
    token = add_draft(store, A_CHANGE, "summary")
    use_case = UpdateEvent(store, writer)

    with pytest.raises(GraphRequestError):
        use_case.execute(token)
    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert writer.calls == 1
