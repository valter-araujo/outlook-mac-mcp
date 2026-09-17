from dataclasses import dataclass

import pytest

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.domain.errors import DraftNotFoundError


@dataclass(frozen=True, slots=True)
class ADraft:
    token: str
    payload: str


def store_with(payload: str) -> tuple[DraftStore[ADraft], ADraft]:
    store: DraftStore[ADraft] = DraftStore()
    draft = store.add(lambda token: ADraft(token=token, payload=payload))
    return store, draft


def test_hands_back_the_built_draft_under_an_opaque_token() -> None:
    _, draft = store_with("summary")

    assert draft.payload == "summary"
    assert len(draft.token) >= 16
    assert "summary" not in draft.token


def test_passes_the_freshly_minted_token_to_the_builder() -> None:
    store: DraftStore[ADraft] = DraftStore()
    seen_tokens: list[str] = []

    def build(token: str) -> ADraft:
        seen_tokens.append(token)
        return ADraft(token=token, payload="x")

    draft = store.add(build)

    assert seen_tokens == [draft.token]


def test_issues_a_different_token_each_time() -> None:
    store: DraftStore[ADraft] = DraftStore()

    first = store.add(lambda token: ADraft(token=token, payload="a"))
    second = store.add(lambda token: ADraft(token=token, payload="b"))

    assert first.token != second.token


def test_take_returns_the_draft_once() -> None:
    store, draft = store_with("summary")

    assert store.take(draft.token) == draft
    with pytest.raises(DraftNotFoundError):
        store.take(draft.token)


def test_take_rejects_a_token_it_never_issued() -> None:
    store: DraftStore[ADraft] = DraftStore()

    with pytest.raises(DraftNotFoundError):
        store.take("not-a-token")


def test_taking_one_draft_leaves_the_others() -> None:
    store: DraftStore[ADraft] = DraftStore()
    first = store.add(lambda token: ADraft(token=token, payload="first"))
    second = store.add(lambda token: ADraft(token=token, payload="second"))

    store.take(first.token)

    assert store.take(second.token) == second


def test_a_token_from_one_store_does_not_exist_in_another() -> None:
    """This is the mechanism that keeps a create token from working for update or
    delete: each flow gets its own store, so a token minted by one is simply absent
    from another's dict, not merely rejected by a type check.
    """
    first_store: DraftStore[ADraft] = DraftStore()
    second_store: DraftStore[ADraft] = DraftStore()
    draft = first_store.add(lambda token: ADraft(token=token, payload="x"))

    with pytest.raises(DraftNotFoundError):
        second_store.take(draft.token)
