import json
from dataclasses import dataclass
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.draft_deleter import DraftDeleter
from outlook_mac_mcp.domain.errors import EmailNotFoundError, NotADraftError
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.draft_deleter import GraphDraftDeleter

MESSAGE_URL = f"{GRAPH_BASE_URL}/me/messages/AAMkDRAFT"
DRAFTS_URL = f"{GRAPH_BASE_URL}/me/mailFolders/drafts"


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def deleter() -> GraphDraftDeleter:
    return GraphDraftDeleter(GraphClient(FakeTokenProvider()))


def a_message(*, is_draft: bool = True, folder: str = "DRAFTS-ID") -> dict[str, Any]:
    return {
        "id": "AAMkDRAFT",
        "subject": "Re: Next steps",
        "toRecipients": [{"emailAddress": {"address": "ana@example.com"}}],
        "lastModifiedDateTime": "2026-11-17T19:02:55Z",
        "isDraft": is_draft,
        "parentFolderId": folder,
    }


def mock_lookup(message: dict[str, Any]) -> None:
    respx.get(MESSAGE_URL).mock(return_value=httpx.Response(200, json=message))
    respx.get(DRAFTS_URL).mock(return_value=httpx.Response(200, json={"id": "DRAFTS-ID"}))


def test_satisfies_the_draft_deleter_port(deleter: GraphDraftDeleter) -> None:
    port: DraftDeleter = deleter

    assert port is deleter


@respx.mock
def test_returns_a_draft_that_is_in_drafts(deleter: GraphDraftDeleter) -> None:
    mock_lookup(a_message())

    draft = deleter.get_draft("AAMkDRAFT")

    assert (draft.subject, [r.address for r in draft.to]) == (
        "Re: Next steps",
        ["ana@example.com"],
    )


@respx.mock
def test_refuses_a_message_in_another_folder(deleter: GraphDraftDeleter) -> None:
    mock_lookup(a_message(folder="INBOX-ID"))

    with pytest.raises(NotADraftError):
        deleter.get_draft("AAMkDRAFT")


@respx.mock
def test_refuses_a_sent_message_moved_into_drafts(deleter: GraphDraftDeleter) -> None:
    mock_lookup(a_message(is_draft=False))

    with pytest.raises(NotADraftError):
        deleter.get_draft("AAMkDRAFT")


@respx.mock
def test_reports_a_missing_message_as_not_found(deleter: GraphDraftDeleter) -> None:
    respx.get(MESSAGE_URL).mock(
        return_value=httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})
    )

    with pytest.raises(EmailNotFoundError):
        deleter.get_draft("AAMkDRAFT")


@respx.mock
def test_delete_moves_the_draft_to_deleted_items_never_http_delete(
    deleter: GraphDraftDeleter,
) -> None:
    """A plain DELETE skipped Deleted Items on a personal account (live, 2026-11-19)."""
    route = respx.post(f"{MESSAGE_URL}/move").mock(
        return_value=httpx.Response(201, json={"id": "AAMkMOVED"})
    )

    deleter.delete("AAMkDRAFT")

    assert json.loads(route.calls.last.request.content) == {"destinationId": "deleteditems"}
    assert [call.request.method for call in respx.calls] == ["POST"]


@respx.mock
def test_delete_reports_a_missing_message_as_not_found(deleter: GraphDraftDeleter) -> None:
    respx.post(f"{MESSAGE_URL}/move").mock(
        return_value=httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})
    )

    with pytest.raises(EmailNotFoundError):
        deleter.delete("AAMkDRAFT")
