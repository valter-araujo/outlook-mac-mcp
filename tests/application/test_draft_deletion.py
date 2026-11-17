from datetime import UTC, datetime

import pytest

from outlook_mac_mcp.application.delete_draft import DeleteDraft
from outlook_mac_mcp.application.draft_deletion_preview import DraftDeletionPreview
from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.preview_draft_deletion import (
    DraftDeletionRequest,
    PreviewDraftDeletion,
)
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import (
    DraftNotFoundError,
    EmailNotFoundError,
    InvalidRequestError,
    NotADraftError,
)
from outlook_mac_mcp.domain.saved_draft import SavedDraft
from tests.fakes.in_memory_draft_deleter import InMemoryDraftDeleter

A_DRAFT = SavedDraft(
    id="AAMkDRAFT",
    subject="Re: Next steps",
    to=(EmailAddress(address="ana@example.com"),),
    last_modified_at=datetime(2026, 11, 17, 19, 2, tzinfo=UTC),
)


def pair() -> tuple[PreviewDraftDeletion, DeleteDraft, InMemoryDraftDeleter]:
    deleter = InMemoryDraftDeleter(drafts=(A_DRAFT,), other_ids=frozenset({"AAMkINBOX"}))
    store: DraftStore[DraftDeletionPreview] = DraftStore()
    return PreviewDraftDeletion(deleter, store), DeleteDraft(store, deleter), deleter


def test_preview_deletes_nothing() -> None:
    preview, _, deleter = pair()

    preview.execute(DraftDeletionRequest(email_id="AAMkDRAFT"))

    assert deleter.deleted == []


def test_preview_summary_identifies_the_draft() -> None:
    preview, _, _ = pair()

    summary = preview.execute(DraftDeletionRequest(email_id="AAMkDRAFT")).summary

    assert summary.startswith("Re: Next steps; to ana@example.com; last edited ")
    assert "Deleted Items" in summary


def test_preview_refuses_a_message_outside_drafts() -> None:
    preview, _, _ = pair()

    with pytest.raises(NotADraftError):
        preview.execute(DraftDeletionRequest(email_id="AAMkINBOX"))


def test_preview_refuses_an_unknown_message() -> None:
    preview, _, _ = pair()

    with pytest.raises(EmailNotFoundError):
        preview.execute(DraftDeletionRequest(email_id="AAMk-missing"))


def test_request_rejects_an_empty_id() -> None:
    with pytest.raises(InvalidRequestError, match="email_id"):
        DraftDeletionRequest(email_id="")


def test_delete_removes_the_previewed_draft() -> None:
    preview, delete, deleter = pair()
    token = preview.execute(DraftDeletionRequest(email_id="AAMkDRAFT")).token

    assert delete.execute(token) == "AAMkDRAFT"
    assert deleter.deleted == ["AAMkDRAFT"]


def test_delete_refuses_a_reused_token() -> None:
    preview, delete, deleter = pair()
    token = preview.execute(DraftDeletionRequest(email_id="AAMkDRAFT")).token
    delete.execute(token)

    with pytest.raises(DraftNotFoundError):
        delete.execute(token)

    assert deleter.deleted == ["AAMkDRAFT"]


def test_delete_refuses_a_draft_moved_out_of_drafts_since_the_preview() -> None:
    preview, delete, deleter = pair()
    token = preview.execute(DraftDeletionRequest(email_id="AAMkDRAFT")).token
    deleter.move_out_of_drafts("AAMkDRAFT")

    with pytest.raises(NotADraftError):
        delete.execute(token)

    assert deleter.deleted == []
