from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.errors import EmailNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError
from outlook_mac_mcp.infrastructure.graph.mail_repository import (
    MALFORMED_ID_ERROR_CODE,
    GraphMailRepository,
)

INBOX_URL = f"{GRAPH_BASE_URL}/me/mailFolders/inbox/messages"
ARCHIVE_URL = f"{GRAPH_BASE_URL}/me/mailFolders/archive/messages"
MESSAGE_URL = f"{GRAPH_BASE_URL}/me/messages"
AN_ID = "AAMkAGI2"


def graph_message(message_id: str, *, received_at: str = "2026-09-14T12:30:00Z") -> dict[str, Any]:
    return {
        "id": message_id,
        "subject": "Quarterly review",
        "from": {"emailAddress": {"name": "Ana Lima", "address": "ana@example.com"}},
        "receivedDateTime": received_at,
        "isRead": False,
        "hasAttachments": False,
        "bodyPreview": "preview",
    }


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def repository() -> GraphMailRepository:
    return GraphMailRepository(GraphClient(FakeTokenProvider()))


def test_satisfies_the_mail_repository_port(repository: GraphMailRepository) -> None:
    port: MailRepository = repository

    assert port is repository


@respx.mock
def test_asks_graph_for_unread_messages_newest_first(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.list_unread(FolderName.INBOX, limit=20)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$filter"] == "isRead eq false"
    assert parameters["$orderby"] == "receivedDateTime desc"
    assert parameters["$top"] == "20"


@respx.mock
def test_passes_the_limit_through_as_the_page_size(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.list_unread(FolderName.INBOX, limit=3)

    assert dict(route.calls.last.request.url.params)["$top"] == "3"


@respx.mock
def test_requests_only_the_fields_the_domain_needs(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.list_unread(FolderName.INBOX, limit=20)

    selected = dict(route.calls.last.request.url.params)["$select"].split(",")
    assert "bodyPreview" in selected
    assert "body" not in selected


@respx.mock
def test_reads_the_requested_folder(repository: GraphMailRepository) -> None:
    archive = respx.get(ARCHIVE_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.list_unread(FolderName.ARCHIVE, limit=20)

    assert archive.call_count == 1


@respx.mock
def test_maps_every_returned_message_in_the_order_graph_gave_them(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    graph_message("newest", received_at="2026-09-14T12:30:00Z"),
                    graph_message("oldest", received_at="2026-09-14T08:00:00Z"),
                ]
            },
        )
    )

    emails = repository.list_unread(FolderName.INBOX, limit=20)

    assert [email.id for email in emails] == ["newest", "oldest"]
    assert emails[0].received_at == datetime(2026, 9, 14, 12, 30, tzinfo=UTC)


@respx.mock
def test_returns_empty_when_the_folder_has_no_unread(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    assert repository.list_unread(FolderName.INBOX, limit=20) == ()


@respx.mock
def test_raises_when_graph_rejects_the_filter_and_sort_combination(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(400, json={"error": {"code": "InefficientFilter"}})
    )

    with pytest.raises(GraphRequestError, match="InefficientFilter"):
        repository.list_unread(FolderName.INBOX, limit=20)


@respx.mock
def test_raises_when_the_collection_has_no_value_array(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"error": None}))

    with pytest.raises(GraphResponseError):
        repository.list_unread(FolderName.INBOX, limit=20)


@respx.mock
def test_raises_when_the_collection_holds_something_that_is_not_a_message(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": ["not-a-message"]}))

    with pytest.raises(GraphResponseError):
        repository.list_unread(FolderName.INBOX, limit=20)


@respx.mock
def test_asks_graph_for_the_body_as_plain_text(repository: GraphMailRepository) -> None:
    route = respx.get(f"{MESSAGE_URL}/{AN_ID}").mock(
        return_value=httpx.Response(200, json=graph_message(AN_ID))
    )

    repository.get_by_id(AN_ID)

    assert route.calls.last.request.headers["Prefer"] == 'outlook.body-content-type="text"'


@respx.mock
def test_selects_the_body_alongside_the_listed_fields(repository: GraphMailRepository) -> None:
    route = respx.get(f"{MESSAGE_URL}/{AN_ID}").mock(
        return_value=httpx.Response(200, json=graph_message(AN_ID))
    )

    repository.get_by_id(AN_ID)

    assert "body" in dict(route.calls.last.request.url.params)["$select"].split(",")


@respx.mock
def test_returns_the_email_with_its_body(repository: GraphMailRepository) -> None:
    message = graph_message(AN_ID) | {"body": {"contentType": "text", "content": "The full text."}}
    respx.get(f"{MESSAGE_URL}/{AN_ID}").mock(return_value=httpx.Response(200, json=message))

    detail = repository.get_by_id(AN_ID)

    assert detail.email.id == AN_ID
    assert detail.body == "The full text."


@respx.mock
def test_percent_encodes_an_id_that_carries_url_characters(
    repository: GraphMailRepository,
) -> None:
    route = respx.get(f"{MESSAGE_URL}/AAMk%2FGI%2B2%3D").mock(
        return_value=httpx.Response(200, json=graph_message("AAMk/GI+2="))
    )

    repository.get_by_id("AAMk/GI+2=")

    assert route.call_count == 1


@respx.mock
def test_an_id_that_looks_like_traversal_cannot_change_the_path(
    repository: GraphMailRepository,
) -> None:
    elsewhere = respx.get(f"{GRAPH_BASE_URL}/me/mailFolders").mock(
        return_value=httpx.Response(200, json={})
    )
    respx.get(f"{MESSAGE_URL}/..%2F..%2FmailFolders").mock(
        return_value=httpx.Response(200, json=graph_message("x"))
    )

    repository.get_by_id("../../mailFolders")

    assert elsewhere.call_count == 0


@respx.mock
def test_raises_email_not_found_when_graph_answers_404(repository: GraphMailRepository) -> None:
    respx.get(f"{MESSAGE_URL}/{AN_ID}").mock(
        return_value=httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})
    )

    with pytest.raises(EmailNotFoundError):
        repository.get_by_id(AN_ID)


@respx.mock
def test_does_not_disguise_other_failures_as_not_found(
    repository: GraphMailRepository,
) -> None:
    respx.get(f"{MESSAGE_URL}/{AN_ID}").mock(
        return_value=httpx.Response(403, json={"error": {"code": "ErrorAccessDenied"}})
    )

    with pytest.raises(GraphRequestError):
        repository.get_by_id(AN_ID)


@respx.mock
def test_raises_invalid_request_when_graph_rejects_the_id_as_malformed(
    repository: GraphMailRepository,
) -> None:
    respx.get(f"{MESSAGE_URL}/nope").mock(
        return_value=httpx.Response(400, json={"error": {"code": MALFORMED_ID_ERROR_CODE}})
    )

    with pytest.raises(InvalidRequestError, match="malformed email id"):
        repository.get_by_id("nope")


@respx.mock
def test_matches_the_malformed_id_on_the_code_not_the_message(
    repository: GraphMailRepository,
) -> None:
    """The same code with different wording must still map, and other 400s must not."""
    respx.get(f"{MESSAGE_URL}/nope").mock(
        return_value=httpx.Response(
            400,
            json={"error": {"code": MALFORMED_ID_ERROR_CODE, "message": "reworded by Microsoft"}},
        )
    )

    with pytest.raises(InvalidRequestError):
        repository.get_by_id("nope")


@respx.mock
def test_leaves_other_bad_requests_as_graph_errors(repository: GraphMailRepository) -> None:
    respx.get(f"{MESSAGE_URL}/{AN_ID}").mock(
        return_value=httpx.Response(400, json={"error": {"code": "ErrorInvalidParameter"}})
    )

    with pytest.raises(GraphRequestError):
        repository.get_by_id(AN_ID)


@respx.mock
def test_search_sends_the_term_as_a_quoted_phrase(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(
        SearchEmailsRequest(term="quarterly review", folder=FolderName.INBOX, limit=20)
    )

    assert dict(route.calls.last.request.url.params)["$search"] == '"quarterly review"'


@respx.mock
def test_search_never_sends_an_orderby(repository: GraphMailRepository) -> None:
    """Graph rejects $orderby alongside $search, so relevance order is not negotiable."""
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="deck", folder=FolderName.INBOX, limit=20))

    assert "$orderby" not in dict(route.calls.last.request.url.params)


@respx.mock
def test_search_bounds_the_page_with_top(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="deck", folder=FolderName.INBOX, limit=7))

    assert dict(route.calls.last.request.url.params)["$top"] == "7"


@respx.mock
def test_search_reads_the_requested_folder(repository: GraphMailRepository) -> None:
    archive = respx.get(ARCHIVE_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="deck", folder=FolderName.ARCHIVE, limit=20))

    assert archive.call_count == 1


@respx.mock
def test_search_sends_an_operator_like_term_as_literal_text(
    repository: GraphMailRepository,
) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(
        SearchEmailsRequest(term="deck AND from:ceo@example.com", folder=FolderName.INBOX, limit=20)
    )

    sent = dict(route.calls.last.request.url.params)["$search"]
    assert sent == '"deck AND from:ceo@example.com"'


@respx.mock
def test_search_escapes_a_term_that_tries_to_close_the_phrase(
    repository: GraphMailRepository,
) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(
        SearchEmailsRequest(
            term='deck" AND from:ceo@example.com "', folder=FolderName.INBOX, limit=20
        )
    )

    sent = dict(route.calls.last.request.url.params)["$search"]
    assert sent == '"deck\\" AND from:ceo@example.com \\""'


@respx.mock
def test_search_maps_the_results(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("found")]})
    )

    result = repository.search(SearchEmailsRequest(term="deck", folder=FolderName.INBOX, limit=20))

    assert [email.id for email in result] == ["found"]


@respx.mock
def test_search_returns_empty_when_nothing_matches(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    assert (
        repository.search(SearchEmailsRequest(term="deck", folder=FolderName.INBOX, limit=20)) == ()
    )


@respx.mock
def test_search_sends_no_restriction_for_the_any_scope(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso", scope=SearchScope.ANY))

    assert dict(route.calls.last.request.url.params)["$search"] == '"Contoso"'


@respx.mock
def test_search_sends_a_subject_restriction(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso", scope=SearchScope.SUBJECT))

    assert dict(route.calls.last.request.url.params)["$search"] == 'subject:"Contoso"'


@respx.mock
def test_search_sends_a_sender_restriction(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso", scope=SearchScope.SENDER))

    assert dict(route.calls.last.request.url.params)["$search"] == 'from:"Contoso"'


@respx.mock
def test_search_defaults_to_the_any_scope(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso"))

    assert dict(route.calls.last.request.url.params)["$search"] == '"Contoso"'


@respx.mock
def test_a_scoped_search_still_escapes_a_breakout_term(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(
        SearchEmailsRequest(term='x" OR from:ceo@example.com "', scope=SearchScope.SUBJECT)
    )

    sent = dict(route.calls.last.request.url.params)["$search"]
    assert sent == 'subject:"x\\" OR from:ceo@example.com \\""'
