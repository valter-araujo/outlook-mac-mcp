import pytest

from outlook_mac_mcp.application.create_reply_draft import CreateReplyDraft
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.preview_reply_draft import PreviewReplyDraft, ReplyDraftRequest
from outlook_mac_mcp.application.reply_draft_preview import ReplyDraftPreview
from outlook_mac_mcp.domain.errors import DraftNotFoundError
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from tests.application.reply_fixtures import ME, THREAD_MESSAGE
from tests.fakes.in_memory_reply_draft_writer import InMemoryReplyDraftWriter

A_REQUEST = ReplyDraftRequest(email_id="AAMk1", mode=ReplyMode.REPLY_ALL, body="Thanks.")


def pair() -> tuple[PreviewReplyDraft, CreateReplyDraft, InMemoryReplyDraftWriter]:
    writer = InMemoryReplyDraftWriter(sources=(THREAD_MESSAGE,), own_addresses=frozenset({ME}))
    store: DraftStore[ReplyDraftPreview] = DraftStore()
    return PreviewReplyDraft(writer, store), CreateReplyDraft(store, writer), writer


def test_saves_exactly_the_previewed_reply() -> None:
    preview, create, writer = pair()
    previewed = preview.execute(A_REQUEST)

    create.execute(previewed.token)

    assert writer.created == [previewed.draft]


def test_returns_the_saved_drafts_id() -> None:
    preview, create, _ = pair()

    created = create.execute(preview.execute(A_REQUEST).token)

    assert created.id == "reply-1"


def test_refuses_a_token_that_was_already_used() -> None:
    preview, create, writer = pair()
    token = preview.execute(A_REQUEST).token
    create.execute(token)

    with pytest.raises(DraftNotFoundError):
        create.execute(token)

    assert len(writer.created) == 1


def test_refuses_a_token_it_never_issued() -> None:
    _, create, writer = pair()

    with pytest.raises(DraftNotFoundError):
        create.execute("not-a-token")

    assert writer.created == []
