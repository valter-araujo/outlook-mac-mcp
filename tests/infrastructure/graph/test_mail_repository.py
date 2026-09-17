from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any

import httpx
import pytest
import respx

from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.ports.mail_repository import MailRepository
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.email_search_page import EmailSearchPage
from outlook_mac_mcp.domain.errors import EmailNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.domain.sort_order import SortOrder
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import (
    GraphRequestError,
    GraphResponseError,
    UnsupportedHostError,
)
from outlook_mac_mcp.infrastructure.graph.mail_repository import (
    MALFORMED_ID_ERROR_CODE,
    GraphMailRepository,
)
from outlook_mac_mcp.infrastructure.graph.search_match_count import COUNT_CEILING

INBOX_URL = f"{GRAPH_BASE_URL}/me/mailFolders/inbox/messages"
ARCHIVE_URL = f"{GRAPH_BASE_URL}/me/mailFolders/archive/messages"
MESSAGE_URL = f"{GRAPH_BASE_URL}/me/messages"
AN_ID = "AAMkAGI2"
NOTHING_UNREAD = {"value": [], "@odata.count": 0}
A_SEARCH_REQUEST = SearchEmailsRequest(term="deck", folders=(FolderName.INBOX,), limit=20)


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
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_unread((FolderName.INBOX,), limit=20)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$filter"] == "isRead eq false"
    assert parameters["$orderby"] == "receivedDateTime desc"
    assert parameters["$top"] == "20"


@respx.mock
def test_passes_the_limit_through_as_the_page_size(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_unread((FolderName.INBOX,), limit=3)

    assert dict(route.calls.last.request.url.params)["$top"] == "3"


@respx.mock
def test_requests_only_the_fields_the_domain_needs(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_unread((FolderName.INBOX,), limit=20)

    selected = dict(route.calls.last.request.url.params)["$select"].split(",")
    assert "bodyPreview" in selected
    assert "body" not in selected


@respx.mock
def test_reads_the_requested_folder(repository: GraphMailRepository) -> None:
    archive = respx.get(ARCHIVE_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_unread((FolderName.ARCHIVE,), limit=20)

    assert archive.call_count == 1


@respx.mock
def test_more_than_one_folder_merges_newest_first_and_sums_the_totals(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [graph_message("inbox-one", received_at="2026-09-14T08:00:00Z")],
                "@odata.count": 5,
            },
        )
    )
    respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [graph_message("archive-one", received_at="2026-09-14T12:00:00Z")],
                "@odata.count": 3,
            },
        )
    )

    page = repository.list_unread((FolderName.INBOX, FolderName.ARCHIVE), limit=20)

    assert [email.id for email in page.items] == ["archive-one", "inbox-one"]
    assert page.total == 8
    assert page.total_is_exact is True


@respx.mock
def test_more_than_one_folder_truncates_the_merged_list_to_the_limit(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [graph_message("inbox-one", received_at="2026-09-14T12:00:00Z")],
                "@odata.count": 1,
            },
        )
    )
    respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [graph_message("archive-one", received_at="2026-09-14T08:00:00Z")],
                "@odata.count": 1,
            },
        )
    )

    page = repository.list_unread((FolderName.INBOX, FolderName.ARCHIVE), limit=1)

    assert [email.id for email in page.items] == ["inbox-one"]
    assert page.total == 2


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
                ],
                "@odata.count": 2,
            },
        )
    )

    emails = repository.list_unread((FolderName.INBOX,), limit=20).items

    assert [email.id for email in emails] == ["newest", "oldest"]
    assert emails[0].received_at == datetime(2026, 9, 14, 12, 30, tzinfo=UTC)


@respx.mock
def test_returns_empty_when_the_folder_has_no_unread(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    page = repository.list_unread((FolderName.INBOX,), limit=20)

    assert page == Page(items=(), total=0, total_is_exact=True)


@respx.mock
def test_asks_graph_to_count_the_unread_in_the_same_request(
    repository: GraphMailRepository,
) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_unread((FolderName.INBOX,), limit=20)

    assert dict(route.calls.last.request.url.params)["$count"] == "true"
    assert route.call_count == 1


@respx.mock
def test_reports_the_folder_count_as_an_exact_total(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200, json={"value": [graph_message("one"), graph_message("two")], "@odata.count": 57}
        )
    )

    page = repository.list_unread((FolderName.INBOX,), limit=2)

    assert len(page.items) == 2
    assert page.total == 57
    assert page.total_is_exact is True


@respx.mock
def test_raises_when_the_count_is_missing(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    with pytest.raises(GraphResponseError):
        repository.list_unread((FolderName.INBOX,), limit=20)


@respx.mock
def test_raises_when_the_count_is_not_a_number(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.count": "57"})
    )

    with pytest.raises(GraphResponseError):
        repository.list_unread((FolderName.INBOX,), limit=20)


@respx.mock
def test_raises_when_graph_rejects_the_filter_and_sort_combination(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(400, json={"error": {"code": "InefficientFilter"}})
    )

    with pytest.raises(GraphRequestError, match="InefficientFilter"):
        repository.list_unread((FolderName.INBOX,), limit=20)


@respx.mock
def test_raises_when_the_collection_has_no_value_array(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"error": None}))

    with pytest.raises(GraphResponseError):
        repository.list_unread((FolderName.INBOX,), limit=20)


@respx.mock
def test_raises_when_the_collection_holds_something_that_is_not_a_message(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": ["not-a-message"]}))

    with pytest.raises(GraphResponseError):
        repository.list_unread((FolderName.INBOX,), limit=20)


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
        SearchEmailsRequest(term="quarterly review", folders=(FolderName.INBOX,), limit=20)
    )

    assert dict(route.calls.last.request.url.params)["$search"] == '"\\"quarterly review\\""'


@respx.mock
def test_search_never_sends_an_orderby(repository: GraphMailRepository) -> None:
    """Graph rejects $orderby alongside $search, so relevance order is not negotiable."""
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="deck", folders=(FolderName.INBOX,), limit=20))

    assert "$orderby" not in dict(route.calls.last.request.url.params)


@respx.mock
def test_search_bounds_the_page_with_top(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="deck", folders=(FolderName.INBOX,), limit=7))

    assert dict(route.calls.last.request.url.params)["$top"] == "7"


@respx.mock
def test_search_reads_the_requested_folder(repository: GraphMailRepository) -> None:
    archive = respx.get(ARCHIVE_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="deck", folders=(FolderName.ARCHIVE,), limit=20))

    assert archive.call_count == 1


@respx.mock
def test_search_more_than_one_folder_concatenates_grouped_by_folder(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("inbox-hit")]})
    )
    respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("archive-hit")]})
    )

    page = repository.search(
        SearchEmailsRequest(term="deck", folders=(FolderName.INBOX, FolderName.ARCHIVE), limit=20)
    )

    assert [email.id for email in page.items] == ["inbox-hit", "archive-hit"]
    assert page.total == 2
    assert page.total_is_exact is True
    assert page.next_page_token is None


@respx.mock
def test_search_more_than_one_folder_recounts_across_every_folder(
    repository: GraphMailRepository,
) -> None:
    """Neither folder's own page is the last one, so the cheap shortcut cannot apply:
    the total is the real count, summed across both folders.
    """
    next_link = f"{INBOX_URL}?%24skip=1"
    # The $select=id-specific routes must be registered before the bare ones: respx
    # matches in registration order, and a bare route matches any query to that path,
    # including the count walk's, unless a more specific route was already checked.
    respx.get(INBOX_URL, params__contains={"$select": "id"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "1"}, {"id": "2"}]})
    )
    respx.get(ARCHIVE_URL, params__contains={"$select": "id"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "1"}]})
    )
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200, json={"value": [graph_message("inbox-hit")], "@odata.nextLink": next_link}
        )
    )
    respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("archive-hit")]})
    )

    page = repository.search(
        SearchEmailsRequest(term="deck", folders=(FolderName.INBOX, FolderName.ARCHIVE), limit=20)
    )

    assert page.total == 3
    assert page.total_is_exact is True
    assert page.next_page_token is None


@respx.mock
def test_search_sends_an_operator_like_term_as_literal_text(
    repository: GraphMailRepository,
) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(
        SearchEmailsRequest(
            term="deck AND from:ceo@example.com", folders=(FolderName.INBOX,), limit=20
        )
    )

    sent = dict(route.calls.last.request.url.params)["$search"]
    assert sent == '"\\"deck AND from:ceo@example.com\\""'


@respx.mock
def test_search_refuses_a_term_that_tries_to_close_the_phrase(
    repository: GraphMailRepository,
) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    with pytest.raises(InvalidRequestError):
        repository.search(
            SearchEmailsRequest(
                term='deck" AND from:ceo@example.com "', folders=(FolderName.INBOX,), limit=20
            )
        )

    assert route.call_count == 0


@respx.mock
def test_search_maps_the_results(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("found")]})
    )

    result = repository.search(A_SEARCH_REQUEST)

    assert [email.id for email in result.items] == ["found"]


@respx.mock
def test_search_returns_empty_when_nothing_matches(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    page = repository.search(A_SEARCH_REQUEST)

    assert page == EmailSearchPage(items=(), total=0, total_is_exact=True, next_page_token=None)


@respx.mock
def test_search_reports_an_exact_total_without_counting_when_the_page_is_the_last(
    repository: GraphMailRepository,
) -> None:
    route = respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("a"), graph_message("b")]})
    )

    page = repository.search(A_SEARCH_REQUEST)

    assert page.total == 2
    assert page.total_is_exact is True
    assert page.next_page_token is None
    assert route.call_count == 1


@respx.mock
def test_search_counts_the_matches_by_id_when_the_page_is_not_the_last(
    repository: GraphMailRepository,
) -> None:
    next_link = f"{INBOX_URL}?%24skip=1"
    count = respx.get(INBOX_URL, params__contains={"$select": "id"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": str(n)} for n in range(45)]})
    )
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200, json={"value": [graph_message("a")], "@odata.nextLink": next_link}
        )
    )

    page = repository.search(SearchEmailsRequest(term="deck", limit=1))

    assert [email.id for email in page.items] == ["a"]
    assert page.total == 45
    assert page.total_is_exact is True
    assert page.next_page_token == next_link
    assert dict(count.calls.last.request.url.params)["$search"] == '"\\"deck\\""'


@respx.mock
def test_search_reports_the_ceiling_as_a_lower_bound(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL, params__contains={"$select": "id"}).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [{"id": str(n)} for n in range(COUNT_CEILING)],
                "@odata.nextLink": f"{INBOX_URL}?%24skip={COUNT_CEILING}",
            },
        )
    )
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200,
            json={"value": [graph_message("a")], "@odata.nextLink": f"{INBOX_URL}?%24skip=1"},
        )
    )

    page = repository.search(SearchEmailsRequest(term="deck", limit=1))

    assert page.total == COUNT_CEILING
    assert page.total_is_exact is False


@respx.mock
def test_search_follows_a_page_token_instead_of_rebuilding_the_query(
    repository: GraphMailRepository,
) -> None:
    """No route is registered for a freshly built $search query at all: if the adapter
    rebuilt the query instead of following the token, this request would go unmocked and
    the test would fail on that, not merely report a wrong item.

    Regression test: omitting page_token must behave exactly as it always did -- see
    every other search_* test above, none of which pass one.
    """
    next_link = f"{INBOX_URL}?%24skip=20"
    followed = respx.get(next_link).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("second-page")]})
    )
    respx.get(INBOX_URL, params__contains={"$select": "id"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "1"}, {"id": "2"}]})
    )

    page = repository.search(SearchEmailsRequest(term="deck", limit=20, page_token=next_link))

    assert followed.called
    assert [email.id for email in page.items] == ["second-page"]


@respx.mock
def test_search_recounts_matches_when_a_followed_page_is_itself_the_last(
    repository: GraphMailRepository,
) -> None:
    """A later page's own size must never be trusted as the total: it never accounts for
    the matches already returned on earlier pages, unlike a first page with no next link.
    """
    next_link = f"{INBOX_URL}?%24skip=20"
    respx.get(next_link).mock(
        return_value=httpx.Response(200, json={"value": [graph_message("last-page")]})
    )
    respx.get(INBOX_URL, params__contains={"$select": "id"}).mock(
        return_value=httpx.Response(200, json={"value": [{"id": str(n)} for n in range(25)]})
    )

    page = repository.search(SearchEmailsRequest(term="deck", limit=20, page_token=next_link))

    assert page.total == 25
    assert page.total_is_exact is True
    assert page.next_page_token is None


@respx.mock
def test_search_rejects_a_page_token_with_a_non_graph_host(
    repository: GraphMailRepository,
) -> None:
    with pytest.raises(UnsupportedHostError):
        repository.search(
            SearchEmailsRequest(
                term="deck", page_token="https://evil.example.com/me/messages?%24skip=1"
            )
        )

    assert respx.calls.call_count == 0


@respx.mock
def test_search_sends_no_restriction_for_the_any_scope(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso", scope=SearchScope.ANY))

    assert dict(route.calls.last.request.url.params)["$search"] == '"\\"Contoso\\""'


@respx.mock
def test_search_sends_a_subject_restriction(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso", scope=SearchScope.SUBJECT))

    assert dict(route.calls.last.request.url.params)["$search"] == '"subject:\\"Contoso\\""'


@respx.mock
def test_search_sends_a_sender_restriction(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso", scope=SearchScope.SENDER))

    assert dict(route.calls.last.request.url.params)["$search"] == '"from:\\"Contoso\\""'


@respx.mock
def test_search_defaults_to_the_any_scope(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    repository.search(SearchEmailsRequest(term="Contoso"))

    assert dict(route.calls.last.request.url.params)["$search"] == '"\\"Contoso\\""'


@respx.mock
def test_a_scoped_search_refuses_a_breakout_term(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    with pytest.raises(InvalidRequestError):
        repository.search(
            SearchEmailsRequest(term='x" OR from:ceo@example.com "', scope=SearchScope.SUBJECT)
        )

    assert route.call_count == 0


A_LISTING = ListEmailsRequest(
    filters=EmailFilters(is_read=False, sender="ana@example.com"), sort=SortOrder.OLDEST, limit=7
)


@respx.mock
def test_list_sends_one_filter_an_order_and_a_count(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_matching(A_LISTING)

    parameters = dict(route.calls.last.request.url.params)
    assert parameters["$filter"] == (
        "receivedDateTime ge 1970-01-01T00:00:00Z and isRead eq false and "
        "from/emailAddress/address eq 'ana@example.com'"
    )
    assert parameters["$orderby"] == "receivedDateTime asc"
    assert parameters["$count"] == "true"
    assert parameters["$top"] == "7"
    assert "bodyPreview" in parameters["$select"].split(",")


@respx.mock
def test_list_sends_no_filter_when_nothing_is_restricted(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_matching(ListEmailsRequest())

    parameters = dict(route.calls.last.request.url.params)
    assert "$filter" not in parameters
    assert parameters["$orderby"] == "receivedDateTime desc"


@respx.mock
def test_list_reads_the_requested_folder(repository: GraphMailRepository) -> None:
    archive = respx.get(ARCHIVE_URL).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))

    repository.list_matching(ListEmailsRequest(filters=EmailFilters(folders=(FolderName.ARCHIVE,))))

    assert archive.call_count == 1


@respx.mock
def test_list_maps_the_page_and_reads_the_exact_total(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200, json={"value": [graph_message("one"), graph_message("two")], "@odata.count": 41}
        )
    )

    page = repository.list_matching(ListEmailsRequest(limit=2))

    assert [email.id for email in page.items] == ["one", "two"]
    assert page.total == 41
    assert page.total_is_exact is True


@respx.mock
def test_list_raises_when_the_count_is_missing(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    with pytest.raises(GraphResponseError):
        repository.list_matching(ListEmailsRequest())


@respx.mock
def test_list_more_than_one_folder_merges_and_sums_the_totals(
    repository: GraphMailRepository,
) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [graph_message("inbox-one", received_at="2026-09-14T08:00:00Z")],
                "@odata.count": 10,
            },
        )
    )
    respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [graph_message("archive-one", received_at="2026-09-14T12:00:00Z")],
                "@odata.count": 4,
            },
        )
    )

    page = repository.list_matching(
        ListEmailsRequest(filters=EmailFilters(folders=(FolderName.INBOX, FolderName.ARCHIVE)))
    )

    assert [email.id for email in page.items] == ["archive-one", "inbox-one"]
    assert page.total == 14
    assert page.total_is_exact is True


@respx.mock
def test_count_asks_for_the_number_and_the_smallest_page(repository: GraphMailRepository) -> None:
    route = respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "x"}], "@odata.count": 1204})
    )

    total = repository.count_matching(EmailFilters(is_read=False, sender="ana@example.com"))

    parameters = dict(route.calls.last.request.url.params)
    assert total == 1204
    assert parameters["$filter"] == (
        "isRead eq false and from/emailAddress/address eq 'ana@example.com'"
    )
    assert parameters["$count"] == "true"
    assert parameters["$top"] == "1"
    assert parameters["$select"] == "id"
    assert "$orderby" not in parameters


@respx.mock
def test_count_without_filters_counts_the_whole_folder(repository: GraphMailRepository) -> None:
    route = respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.count": 0})
    )

    assert repository.count_matching(EmailFilters(folders=(FolderName.ARCHIVE,))) == 0
    assert "$filter" not in dict(route.calls.last.request.url.params)


@respx.mock
def test_count_more_than_one_folder_sums_each_exact_count(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.count": 7})
    )
    respx.get(ARCHIVE_URL).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.count": 2})
    )

    total = repository.count_matching(EmailFilters(folders=(FolderName.INBOX, FolderName.ARCHIVE)))

    assert total == 9


@respx.mock
def test_count_raises_when_graph_returns_no_count(repository: GraphMailRepository) -> None:
    respx.get(INBOX_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    with pytest.raises(GraphResponseError):
        repository.count_matching(EmailFilters())


@respx.mock
def test_a_resolved_custom_folder_id_is_used_exactly_as_given(
    repository: GraphMailRepository,
) -> None:
    """A custom folder's `folders` entry is a Graph id resolved elsewhere (list_custom's
    tree walk), never a well-known name -- confirming it reaches the same URL segment a
    FolderName would, with no enum-specific handling standing in the way.
    """
    unread_url = f"{GRAPH_BASE_URL}/me/mailFolders/AAMkAG-custom-unread/messages"
    count_url = f"{GRAPH_BASE_URL}/me/mailFolders/AAMkAG-custom-count/messages"
    unread_route = respx.get(unread_url).mock(return_value=httpx.Response(200, json=NOTHING_UNREAD))
    count_route = respx.get(count_url).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.count": 3})
    )

    repository.list_unread(("AAMkAG-custom-unread",), limit=20)
    total = repository.count_matching(EmailFilters(folders=("AAMkAG-custom-count",)))

    assert unread_route.call_count == 1
    assert count_route.call_count == 1
    assert total == 3
