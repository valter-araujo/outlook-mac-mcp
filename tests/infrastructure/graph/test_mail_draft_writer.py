import json
from dataclasses import dataclass

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.mail_draft_writer import MailDraftWriter
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import MailDraftWriteUnconfirmedError
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
from outlook_mac_mcp.infrastructure.graph.mail_draft_writer import GraphMailDraftWriter

MESSAGES_URL = f"{GRAPH_BASE_URL}/me/messages"
A_DRAFT = NewMailDraft(
    subject="Hello",
    body="Hi Ana.",
    to=(EmailAddress(address="ana@example.com"),),
    cc=(EmailAddress(address="bo@example.com"),),
)


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def writer() -> GraphMailDraftWriter:
    return GraphMailDraftWriter(GraphClient(FakeTokenProvider()))


def test_satisfies_the_mail_draft_writer_port(writer: GraphMailDraftWriter) -> None:
    port: MailDraftWriter = writer

    assert port is writer


@respx.mock
def test_posts_the_draft_as_a_plain_text_message(writer: GraphMailDraftWriter) -> None:
    route = respx.post(MESSAGES_URL).mock(
        return_value=httpx.Response(201, json={"id": "AAMkDRAFT", "isDraft": True})
    )

    writer.create(A_DRAFT)

    assert json.loads(route.calls.last.request.content) == {
        "subject": "Hello",
        "body": {"contentType": "text", "content": "Hi Ana."},
        "toRecipients": [{"emailAddress": {"address": "ana@example.com"}}],
        "ccRecipients": [{"emailAddress": {"address": "bo@example.com"}}],
    }


@respx.mock
def test_returns_the_id_graph_assigned(writer: GraphMailDraftWriter) -> None:
    respx.post(MESSAGES_URL).mock(return_value=httpx.Response(201, json={"id": "AAMkDRAFT"}))

    assert writer.create(A_DRAFT).id == "AAMkDRAFT"


@respx.mock
def test_makes_exactly_one_request_and_never_touches_a_send_endpoint(
    writer: GraphMailDraftWriter,
) -> None:
    route = respx.post(MESSAGES_URL).mock(
        return_value=httpx.Response(201, json={"id": "AAMkDRAFT"})
    )

    writer.create(A_DRAFT)

    assert [call.request.url.path for call in respx.calls] == ["/v1.0/me/messages"]
    assert route.call_count == 1


@respx.mock
def test_reports_unconfirmed_when_the_saved_draft_has_no_id(writer: GraphMailDraftWriter) -> None:
    respx.post(MESSAGES_URL).mock(return_value=httpx.Response(201, json={"isDraft": True}))

    with pytest.raises(MailDraftWriteUnconfirmedError, match="Drafts folder"):
        writer.create(A_DRAFT)


@respx.mock
def test_lets_a_refused_request_surface_as_a_graph_error(writer: GraphMailDraftWriter) -> None:
    respx.post(MESSAGES_URL).mock(
        return_value=httpx.Response(
            403, json={"error": {"code": "ErrorAccessDenied", "message": "denied"}}
        )
    )

    with pytest.raises(GraphRequestError) as raised:
        writer.create(A_DRAFT)

    assert raised.value.status_code == 403
