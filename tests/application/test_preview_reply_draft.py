import pytest

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.preview_reply_draft import PreviewReplyDraft, ReplyDraftRequest
from outlook_mac_mcp.application.reply_draft_preview import ReplyDraftPreview
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import EmailNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from tests.application.reply_fixtures import (
    ME,
    NO_REPLY_MESSAGE,
    REDIRECTED_MESSAGE,
    THREAD_MESSAGE,
)
from tests.fakes.in_memory_reply_draft_writer import InMemoryReplyDraftWriter


def preview_with(writer: InMemoryReplyDraftWriter) -> PreviewReplyDraft:
    store: DraftStore[ReplyDraftPreview] = DraftStore()
    return PreviewReplyDraft(writer, store)


def a_writer() -> InMemoryReplyDraftWriter:
    return InMemoryReplyDraftWriter(
        sources=(THREAD_MESSAGE, NO_REPLY_MESSAGE, REDIRECTED_MESSAGE),
        own_addresses=frozenset({ME}),
    )


def request(email_id: str = "AAMk1", mode: ReplyMode = ReplyMode.REPLY_ALL) -> ReplyDraftRequest:
    return ReplyDraftRequest(email_id=email_id, mode=mode, body="Thanks, Tuesday works.")


def test_saves_nothing() -> None:
    writer = a_writer()

    preview_with(writer).execute(request())

    assert writer.created == []


def test_reply_all_resolves_every_recipient_but_the_user() -> None:
    preview = preview_with(a_writer()).execute(request())

    recipients = preview.draft.recipients
    assert [r.address for r in recipients.to] == ["ana@example.com", "bo@example.com"]
    assert [r.address for r in recipients.cc] == ["cy@example.com", "di@example.com"]


def test_summary_lists_the_reply_subject_and_every_recipient() -> None:
    preview = preview_with(a_writer()).execute(request())

    assert preview.summary.startswith("Re: Next steps; reply all; ")
    assert "to ana@example.com, bo@example.com" in preview.summary
    assert "cc cy@example.com, di@example.com" in preview.summary
    assert ME not in preview.summary


def test_summary_does_not_double_an_existing_reply_prefix() -> None:
    preview = preview_with(a_writer()).execute(request("AAMk2"))

    assert preview.summary.startswith("Re: Your application;")


def test_warns_when_a_recipient_looks_like_a_no_reply_address() -> None:
    preview = preview_with(a_writer()).execute(request("AAMk2"))

    assert any("recruiting@email.example.com" in warning for warning in preview.warnings)


def test_warns_when_reply_to_redirects_the_reply() -> None:
    preview = preview_with(a_writer()).execute(request("AAMk3", ReplyMode.REPLY))

    assert [r.address for r in preview.draft.recipients.to] == ["recruiter@example.com"]
    assert any("instead of its sender ats@example.com" in w for w in preview.warnings)


def test_has_no_warnings_for_an_ordinary_thread() -> None:
    preview = preview_with(a_writer()).execute(request())

    assert preview.warnings == ()


def test_adds_the_extra_recipients() -> None:
    extra = ReplyDraftRequest(
        email_id="AAMk1",
        mode=ReplyMode.REPLY,
        body="Hi",
        extra_cc=(EmailAddress(address="eve@example.com"),),
    )

    preview = preview_with(a_writer()).execute(extra)

    assert [r.address for r in preview.draft.recipients.cc] == ["eve@example.com"]


def test_refuses_an_unknown_message() -> None:
    with pytest.raises(EmailNotFoundError):
        preview_with(a_writer()).execute(request("AAMk-missing"))


def test_rejects_an_empty_email_id() -> None:
    with pytest.raises(InvalidRequestError, match="email_id"):
        ReplyDraftRequest(email_id="", mode=ReplyMode.REPLY, body="Hi")
