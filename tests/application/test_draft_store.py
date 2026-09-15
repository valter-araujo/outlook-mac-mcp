from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.domain.errors import DraftNotFoundError
from outlook_mac_mcp.domain.new_event import NewEvent

NINE = datetime(2026, 9, 15, 9, tzinfo=ZoneInfo("America/Sao_Paulo"))
AN_EVENT = NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1))


def test_hands_back_the_event_under_an_opaque_token() -> None:
    draft = DraftStore().add(AN_EVENT, "summary")

    assert draft.new_event == AN_EVENT
    assert draft.summary == "summary"
    assert len(draft.token) >= 16
    assert "Planning" not in draft.token


def test_issues_a_different_token_each_time() -> None:
    store = DraftStore()

    assert store.add(AN_EVENT, "a").token != store.add(AN_EVENT, "b").token


def test_take_returns_the_draft_once() -> None:
    store = DraftStore()
    draft = store.add(AN_EVENT, "summary")

    assert store.take(draft.token) == draft
    with pytest.raises(DraftNotFoundError):
        store.take(draft.token)


def test_take_rejects_a_token_it_never_issued() -> None:
    with pytest.raises(DraftNotFoundError):
        DraftStore().take("not-a-token")


def test_taking_one_draft_leaves_the_others() -> None:
    store = DraftStore()
    first = store.add(AN_EVENT, "first")
    second = store.add(AN_EVENT, "second")

    store.take(first.token)

    assert store.take(second.token) == second
