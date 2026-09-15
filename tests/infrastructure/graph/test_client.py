import json
from dataclasses import dataclass

import httpx
import pytest
import respx

from outlook_mac_mcp.infrastructure.graph.client import (
    GRAPH_BASE_URL,
    UNKNOWN_ERROR_CODE,
    GraphClient,
)
from outlook_mac_mcp.infrastructure.graph.errors import (
    GraphRequestError,
    GraphResponseError,
    UnsupportedHostError,
)

MESSAGES_PATH = "/me/messages"
MESSAGES_URL = f"{GRAPH_BASE_URL}{MESSAGES_PATH}"
ANOTHER_HOST_URL = "https://graph.microsoft.com.evil.example/v1.0/me/messages"
SECRET_BODY = {"error": {"code": "ErrorAccessDenied", "message": "mailbox of ceo@example.com"}}


@dataclass
class FakeTokenProvider:
    token: str = "a-token"

    def get_access_token(self) -> str:
        return self.token


@pytest.fixture
def client() -> GraphClient:
    return GraphClient(FakeTokenProvider())


@respx.mock
def test_sends_the_bearer_token_from_the_provider(client: GraphClient) -> None:
    route = respx.get(MESSAGES_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    client.get(MESSAGES_PATH, {})

    assert route.calls.last.request.headers["Authorization"] == "Bearer a-token"


@respx.mock
def test_returns_the_decoded_payload(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(return_value=httpx.Response(200, json={"value": [{"id": "1"}]}))

    payload = client.get(MESSAGES_PATH, {})

    assert payload == {"value": [{"id": "1"}]}


@respx.mock
def test_sends_query_parameters_to_graph(client: GraphClient) -> None:
    route = respx.get(MESSAGES_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    client.get(MESSAGES_PATH, {"$top": 5, "$filter": "isRead eq false"})

    assert dict(route.calls.last.request.url.params) == {"$top": "5", "$filter": "isRead eq false"}


@respx.mock
def test_raises_with_the_status_error_code_and_request_id(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(
        return_value=httpx.Response(403, json=SECRET_BODY, headers={"request-id": "abc-123"})
    )

    with pytest.raises(GraphRequestError) as failure:
        client.get(MESSAGES_PATH, {})

    assert "403" in str(failure.value)
    assert "ErrorAccessDenied" in str(failure.value)
    assert "abc-123" in str(failure.value)


@respx.mock
def test_never_puts_the_error_body_into_the_raised_message(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(return_value=httpx.Response(403, json=SECRET_BODY))

    with pytest.raises(GraphRequestError) as failure:
        client.get(MESSAGES_PATH, {})

    assert "ceo@example.com" not in str(failure.value)


@respx.mock
def test_raises_when_graph_answers_with_something_other_than_json(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(return_value=httpx.Response(200, text="<html>sign in</html>"))

    with pytest.raises(GraphResponseError):
        client.get(MESSAGES_PATH, {})


@respx.mock
def test_raises_when_graph_answers_with_a_json_value_that_is_not_an_object(
    client: GraphClient,
) -> None:
    respx.get(MESSAGES_URL).mock(return_value=httpx.Response(200, json=[1, 2, 3]))

    with pytest.raises(GraphResponseError):
        client.get(MESSAGES_PATH, {})


@respx.mock
def test_refuses_an_absolute_url_without_reaching_the_network(client: GraphClient) -> None:
    route = respx.get(ANOTHER_HOST_URL).mock(return_value=httpx.Response(200, json={}))

    with pytest.raises(UnsupportedHostError):
        client.get(ANOTHER_HOST_URL, {})

    assert route.call_count == 0


@respx.mock
def test_refuses_a_protocol_relative_path_without_reaching_the_network(
    client: GraphClient,
) -> None:
    route = respx.get("https://evil.example/v1.0/me/messages").mock(
        return_value=httpx.Response(200, json={})
    )

    with pytest.raises(UnsupportedHostError):
        client.get("//evil.example/v1.0/me/messages", {})

    assert route.call_count == 0


@respx.mock
def test_does_not_follow_a_redirect_to_another_host(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(
        return_value=httpx.Response(302, headers={"location": ANOTHER_HOST_URL})
    )
    elsewhere = respx.get(ANOTHER_HOST_URL).mock(return_value=httpx.Response(200, json={}))

    with pytest.raises(GraphRequestError):
        client.get(MESSAGES_PATH, {})

    assert elsewhere.call_count == 0


@respx.mock
def test_sends_the_headers_it_is_given(client: GraphClient) -> None:
    route = respx.get(MESSAGES_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    client.get(MESSAGES_PATH, {}, {"Prefer": 'outlook.body-content-type="text"'})

    assert route.calls.last.request.headers["Prefer"] == 'outlook.body-content-type="text"'


@respx.mock
def test_still_sends_the_bearer_token_when_headers_are_given(client: GraphClient) -> None:
    route = respx.get(MESSAGES_URL).mock(return_value=httpx.Response(200, json={"value": []}))

    client.get(MESSAGES_PATH, {}, {"Prefer": "anything"})

    assert route.calls.last.request.headers["Authorization"] == "Bearer a-token"


@respx.mock
def test_reports_the_failing_status_on_the_error(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(return_value=httpx.Response(404, json={}))

    with pytest.raises(GraphRequestError) as failure:
        client.get(MESSAGES_PATH, {})

    assert failure.value.status_code == 404


@respx.mock
def test_reports_the_graph_error_code_on_the_error(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(
        return_value=httpx.Response(400, json={"error": {"code": "ErrorInvalidIdMalformed"}})
    )

    with pytest.raises(GraphRequestError) as failure:
        client.get(MESSAGES_PATH, {})

    assert failure.value.error_code == "ErrorInvalidIdMalformed"


@respx.mock
def test_reports_an_unknown_code_when_the_body_carries_none(client: GraphClient) -> None:
    respx.get(MESSAGES_URL).mock(return_value=httpx.Response(500, text="oops"))

    with pytest.raises(GraphRequestError) as failure:
        client.get(MESSAGES_PATH, {})

    assert failure.value.error_code == UNKNOWN_ERROR_CODE


NEXT_LINK = f"{GRAPH_BASE_URL}/me/messages?%24skiptoken=abc123"


@respx.mock
def test_follows_a_pagination_link(client: GraphClient) -> None:
    respx.get(NEXT_LINK).mock(return_value=httpx.Response(200, json={"value": [{"id": "2"}]}))

    payload = client.follow(NEXT_LINK)

    assert payload == {"value": [{"id": "2"}]}


@respx.mock
def test_sends_the_bearer_token_when_following_a_link(client: GraphClient) -> None:
    route = respx.get(NEXT_LINK).mock(return_value=httpx.Response(200, json={"value": []}))

    client.follow(NEXT_LINK)

    assert route.calls.last.request.headers["Authorization"] == "Bearer a-token"


@respx.mock
def test_refuses_a_pagination_link_to_another_host(client: GraphClient) -> None:
    """A nextLink is server-supplied, so it is exactly the input a host check must survive."""
    route = respx.get(ANOTHER_HOST_URL).mock(return_value=httpx.Response(200, json={}))

    with pytest.raises(UnsupportedHostError):
        client.follow(ANOTHER_HOST_URL)

    assert route.call_count == 0


@respx.mock
def test_refuses_a_plain_http_pagination_link(client: GraphClient) -> None:
    route = respx.get("http://graph.microsoft.com/v1.0/me/messages").mock(
        return_value=httpx.Response(200, json={})
    )

    with pytest.raises(UnsupportedHostError):
        client.follow("http://graph.microsoft.com/v1.0/me/messages")

    assert route.call_count == 0


@respx.mock
def test_raises_when_a_followed_page_fails(client: GraphClient) -> None:
    respx.get(NEXT_LINK).mock(return_value=httpx.Response(429, json={}))

    with pytest.raises(GraphRequestError):
        client.follow(NEXT_LINK)


EVENTS_PATH = "/me/events"
EVENTS_URL = f"{GRAPH_BASE_URL}{EVENTS_PATH}"


@respx.mock
def test_posts_the_body_as_json(client: GraphClient) -> None:
    route = respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json={"id": "new"}))

    payload = client.post(EVENTS_PATH, {"subject": "Planning"})

    request = route.calls.last.request
    assert request.headers["Content-Type"] == "application/json"
    assert json.loads(request.content) == {"subject": "Planning"}
    assert payload == {"id": "new"}


@respx.mock
def test_sends_the_bearer_token_and_extra_headers_when_posting(client: GraphClient) -> None:
    route = respx.post(EVENTS_URL).mock(return_value=httpx.Response(201, json={}))

    client.post(EVENTS_PATH, {}, {"Prefer": 'outlook.timezone="UTC"'})

    request = route.calls.last.request
    assert request.headers["Authorization"] == "Bearer a-token"
    assert request.headers["Prefer"] == 'outlook.timezone="UTC"'


@respx.mock
def test_refuses_to_post_to_an_absolute_url(client: GraphClient) -> None:
    route = respx.post(ANOTHER_HOST_URL).mock(return_value=httpx.Response(201, json={}))

    with pytest.raises(UnsupportedHostError):
        client.post(ANOTHER_HOST_URL, {})

    assert route.call_count == 0


@respx.mock
def test_raises_with_the_error_code_when_a_post_is_rejected(client: GraphClient) -> None:
    respx.post(EVENTS_URL).mock(
        return_value=httpx.Response(400, json={"error": {"code": "ErrorInvalidPropertyRequest"}})
    )

    with pytest.raises(GraphRequestError, match="ErrorInvalidPropertyRequest"):
        client.post(EVENTS_PATH, {})
