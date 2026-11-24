import json
from dataclasses import dataclass

import httpx
import pytest
import respx

from outlook_mac_mcp.application.ports.reply_draft_writer import ReplyDraftWriter
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import (
    EmailNotFoundError,
    InvalidRequestError,
    MailDraftWriteUnconfirmedError,
)
from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import ReplyRecipients
from outlook_mac_mcp.infrastructure.graph.client import GRAPH_BASE_URL, GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError
from outlook_mac_mcp.infrastructure.graph.reply_draft_writer import GraphReplyDraftWriter

MESSAGE_URL = f"{GRAPH_BASE_URL}/me/messages/AAMk%2F1"
ME_URL = f"{GRAPH_BASE_URL}/me"
A_REPLY_ALL = NewReplyDraft(
    source_id="AAMk/1",
    mode=ReplyMode.REPLY_ALL,
    body="Thanks, Tuesday works.",
    recipients=ReplyRecipients(
        to=(EmailAddress(address="ana@example.com"), EmailAddress(address="bo@example.com")),
        cc=(EmailAddress(address="cy@example.com"),),
    ),
)
A_MESSAGE = {
    "id": "AAMk/1",
    "subject": "Next steps",
    "from": {"emailAddress": {"name": "Ana", "address": "ana@example.com"}},
    "replyTo": [{"emailAddress": {"address": "recruiter@example.com"}}],
    "toRecipients": [
        {"emailAddress": {"address": "me@example.com"}},
        {"emailAddress": {"address": "bo@example.com"}},
    ],
    "ccRecipients": [{"emailAddress": {"name": "Cy"}}],
}


@dataclass
class FakeTokenProvider:
    def get_access_token(self) -> str:
        return "a-token"


@pytest.fixture
def writer() -> GraphReplyDraftWriter:
    return GraphReplyDraftWriter(GraphClient(FakeTokenProvider()))


def not_found() -> httpx.Response:
    return httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}})


def test_satisfies_the_reply_draft_writer_port(writer: GraphReplyDraftWriter) -> None:
    port: ReplyDraftWriter = writer

    assert port is writer


@respx.mock
def test_reads_the_source_with_reply_to_and_every_recipient(
    writer: GraphReplyDraftWriter,
) -> None:
    respx.get(MESSAGE_URL).mock(return_value=httpx.Response(200, json=A_MESSAGE))

    source = writer.get_reply_source("AAMk/1")

    assert source.sender == EmailAddress(address="ana@example.com", display_name="Ana")
    assert [r.address for r in source.reply_to] == ["recruiter@example.com"]
    assert [r.address for r in source.to] == ["me@example.com", "bo@example.com"]


@respx.mock
def test_drops_a_recipient_with_no_address(writer: GraphReplyDraftWriter) -> None:
    respx.get(MESSAGE_URL).mock(return_value=httpx.Response(200, json=A_MESSAGE))

    assert writer.get_reply_source("AAMk/1").cc == ()


@respx.mock
def test_selects_only_the_fields_a_reply_needs(writer: GraphReplyDraftWriter) -> None:
    route = respx.get(MESSAGE_URL).mock(return_value=httpx.Response(200, json=A_MESSAGE))

    writer.get_reply_source("AAMk/1")

    assert route.calls.last.request.url.params["$select"] == (
        "id,subject,from,replyTo,toRecipients,ccRecipients"
    )


@respx.mock
def test_reports_a_missing_source_as_not_found(writer: GraphReplyDraftWriter) -> None:
    respx.get(MESSAGE_URL).mock(return_value=not_found())

    with pytest.raises(EmailNotFoundError):
        writer.get_reply_source("AAMk/1")


@respx.mock
def test_reports_a_malformed_id_as_an_invalid_request(writer: GraphReplyDraftWriter) -> None:
    respx.get(MESSAGE_URL).mock(
        return_value=httpx.Response(400, json={"error": {"code": "ErrorInvalidIdMalformed"}})
    )

    with pytest.raises(InvalidRequestError):
        writer.get_reply_source("AAMk/1")


@respx.mock
def test_lets_other_failures_through_unchanged(writer: GraphReplyDraftWriter) -> None:
    respx.get(MESSAGE_URL).mock(return_value=httpx.Response(503, json={}))

    with pytest.raises(GraphRequestError):
        writer.get_reply_source("AAMk/1")


@respx.mock
def test_own_addresses_combine_mail_and_principal_name(writer: GraphReplyDraftWriter) -> None:
    respx.get(ME_URL).mock(
        return_value=httpx.Response(200, json={"mail": None, "userPrincipalName": "Me@Example.com"})
    )

    assert writer.get_own_addresses() == frozenset({"me@example.com"})


@respx.mock
def test_reply_all_posts_to_create_reply_all(writer: GraphReplyDraftWriter) -> None:
    route = respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY", "isDraft": True})
    )

    writer.create_reply(A_REPLY_ALL)

    assert route.call_count == 1


@respx.mock
def test_reply_posts_to_create_reply(writer: GraphReplyDraftWriter) -> None:
    route = respx.post(f"{MESSAGE_URL}/createReply").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY"})
    )

    writer.create_reply(
        NewReplyDraft(
            source_id="AAMk/1",
            mode=ReplyMode.REPLY,
            body="Hi",
            recipients=ReplyRecipients(to=(EmailAddress(address="ana@example.com"),), cc=()),
        )
    )

    assert route.call_count == 1


@respx.mock
def test_sends_the_text_as_comment_and_the_previewed_recipients(
    writer: GraphReplyDraftWriter,
) -> None:
    route = respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY"})
    )

    writer.create_reply(A_REPLY_ALL)

    assert json.loads(route.calls.last.request.content) == {
        "comment": "Thanks, Tuesday works.",
        "message": {
            "toRecipients": [
                {"emailAddress": {"address": "ana@example.com"}},
                {"emailAddress": {"address": "bo@example.com"}},
            ],
            "ccRecipients": [{"emailAddress": {"address": "cy@example.com"}}],
        },
    }


@respx.mock
def test_never_sends_a_message_body_alongside_the_comment(
    writer: GraphReplyDraftWriter,
) -> None:
    """Graph answers 400 when both are present; see reply_draft_payload."""
    route = respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY"})
    )

    writer.create_reply(A_REPLY_ALL)

    assert "body" not in json.loads(route.calls.last.request.content)["message"]


@respx.mock
def test_makes_exactly_one_request_and_never_touches_a_send_endpoint(
    writer: GraphReplyDraftWriter,
) -> None:
    respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY"})
    )

    writer.create_reply(A_REPLY_ALL)

    assert [call.request.url.raw_path.decode() for call in respx.calls] == [
        "/v1.0/me/messages/AAMk%2F1/createReplyAll"
    ]


@respx.mock
def test_returns_the_id_graph_assigned(writer: GraphReplyDraftWriter) -> None:
    respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY"})
    )

    assert writer.create_reply(A_REPLY_ALL).id == "AAMkREPLY"


@respx.mock
def test_reports_unconfirmed_when_the_saved_reply_has_no_id(
    writer: GraphReplyDraftWriter,
) -> None:
    respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"isDraft": True})
    )

    with pytest.raises(MailDraftWriteUnconfirmedError, match="Drafts folder"):
        writer.create_reply(A_REPLY_ALL)


@respx.mock
def test_reports_a_source_deleted_since_the_preview_as_not_found(
    writer: GraphReplyDraftWriter,
) -> None:
    respx.post(f"{MESSAGE_URL}/createReplyAll").mock(return_value=not_found())

    with pytest.raises(EmailNotFoundError):
        writer.create_reply(A_REPLY_ALL)


@respx.mock
def test_escapes_markup_in_the_comment_because_graph_reads_it_as_html(
    writer: GraphReplyDraftWriter,
) -> None:
    """Regression: the live check found an unescaped <b> saved as a live tag."""
    route = respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY"})
    )
    draft = NewReplyDraft(
        source_id="AAMk/1",
        mode=ReplyMode.REPLY_ALL,
        body='<img src="https://tracker.example.com/p.gif"> & <b>hi</b>',
        recipients=A_REPLY_ALL.recipients,
    )

    writer.create_reply(draft)

    assert json.loads(route.calls.last.request.content)["comment"] == (
        "&lt;img src=&quot;https://tracker.example.com/p.gif&quot;&gt; &amp; &lt;b&gt;hi&lt;/b&gt;"
    )


@respx.mock
def test_keeps_line_breaks_in_the_comment(writer: GraphReplyDraftWriter) -> None:
    route = respx.post(f"{MESSAGE_URL}/createReplyAll").mock(
        return_value=httpx.Response(201, json={"id": "AAMkREPLY"})
    )
    draft = NewReplyDraft(
        source_id="AAMk/1",
        mode=ReplyMode.REPLY_ALL,
        body="Hi Ana,\r\nTuesday works.\nBest",
        recipients=A_REPLY_ALL.recipients,
    )

    writer.create_reply(draft)

    assert json.loads(route.calls.last.request.content)["comment"] == (
        "Hi Ana,<br>Tuesday works.<br>Best"
    )
