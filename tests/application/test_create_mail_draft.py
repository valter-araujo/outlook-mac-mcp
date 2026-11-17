import pytest

from outlook_mac_mcp.application.create_mail_draft import CreateMailDraft
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.mail_draft_preview import MailDraftPreview
from outlook_mac_mcp.domain.created_mail_draft import CreatedMailDraft
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import DraftNotFoundError
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
from tests.fakes.in_memory_mail_draft_writer import InMemoryMailDraftWriter

A_DRAFT = NewMailDraft(subject="Hello", body="Hi.", to=(EmailAddress(address="ana@example.com"),))


class FailingMailDraftWriter:
    def __init__(self) -> None:
        self.calls = 0

    def create(self, new_draft: NewMailDraft) -> CreatedMailDraft:
        self.calls += 1
        raise GraphRequestError("Graph returned 503", 503, "ServiceUnavailable")


def add_preview(store: DraftStore[MailDraftPreview]) -> str:
    return store.add(
        lambda token: MailDraftPreview(token=token, summary="summary", new_draft=A_DRAFT)
    ).token


def test_saves_the_previewed_draft_and_returns_its_handle() -> None:
    store: DraftStore[MailDraftPreview] = DraftStore()
    writer = InMemoryMailDraftWriter()

    created = CreateMailDraft(store, writer).execute(add_preview(store))

    assert writer.created == [A_DRAFT]
    assert created.id == "draft-1"


def test_a_token_saves_at_most_one_draft() -> None:
    store: DraftStore[MailDraftPreview] = DraftStore()
    writer = InMemoryMailDraftWriter()
    use_case = CreateMailDraft(store, writer)
    token = add_preview(store)
    use_case.execute(token)

    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert len(writer.created) == 1


def test_an_unknown_token_saves_nothing() -> None:
    writer = InMemoryMailDraftWriter()

    with pytest.raises(DraftNotFoundError):
        CreateMailDraft(DraftStore(), writer).execute("never-issued")

    assert writer.created == []


def test_a_failed_write_consumes_the_preview_so_a_retry_cannot_duplicate() -> None:
    store: DraftStore[MailDraftPreview] = DraftStore()
    writer = FailingMailDraftWriter()
    use_case = CreateMailDraft(store, writer)
    token = add_preview(store)

    with pytest.raises(GraphRequestError):
        use_case.execute(token)
    with pytest.raises(DraftNotFoundError):
        use_case.execute(token)

    assert writer.calls == 1
