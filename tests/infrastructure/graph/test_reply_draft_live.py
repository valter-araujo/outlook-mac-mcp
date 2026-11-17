r"""Live checks of how Graph saves a reply draft built by GraphReplyDraftWriter.

The unit tests only assert the bytes we send. Three properties depend on how Graph
treats them and are checked here against a real mailbox:

1. Escaping, with a positive control: the reply's `comment` carries markup. If Graph
   inserted it as HTML, the saved body would contain a live `<b>` tag; it must instead
   contain the escaped text.
2. Threading and recipients: the draft must be in the original's conversation, and
   carry exactly the To and Cc lines we sent, not a list Graph recomputed.
3. The source message's read state must not change. Reports on Microsoft Q&A say
   createReply can flip a recently modified message to read; this server never
   modifies the source, so it should not happen here.

This WRITES to the mailbox: it saves one reply draft and moves it to Deleted Items
afterwards. Point it at a message you own whose recipients you do not mind appearing
on a draft, through OUTLOOK_MCP_LIVE_REPLY_SOURCE_ID (an id from list_emails).

Not yet measured: run it and record the date and outcome here.

Run with:
    OUTLOOK_MCP_LIVE_REPLY_SOURCE_ID=<id> OUTLOOK_MCP_ENABLE_MAIL_DRAFT=true \
    uv run pytest -m integration tests/infrastructure/graph/test_reply_draft_live.py
"""

import os
import uuid
from collections.abc import Iterator, Mapping
from typing import Any
from urllib.parse import quote

import pytest

from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import resolve_reply_recipients
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.reply_draft_writer import GraphReplyDraftWriter
from outlook_mac_mcp.infrastructure.settings import load_settings

pytestmark = pytest.mark.integration

SOURCE_ID_ENV_VAR = "OUTLOOK_MCP_LIVE_REPLY_SOURCE_ID"
CHECK_FIELDS = "conversationId,isDraft,isRead,body,toRecipients,ccRecipients"


@pytest.fixture
def client() -> GraphClient:
    return GraphClient(DeviceCodeAuthenticator.from_settings(load_settings(os.environ)))


@pytest.fixture
def source_id() -> str:
    value = os.environ.get(SOURCE_ID_ENV_VAR, "")
    if not value:
        pytest.skip(f"set {SOURCE_ID_ENV_VAR} to the id of a message to reply to")
    return value


@pytest.fixture
def saved(client: GraphClient, source_id: str) -> Iterator[tuple[str, str, NewReplyDraft]]:
    writer = GraphReplyDraftWriter(client)
    source = writer.get_reply_source(source_id)
    marker = f"MARKER-{uuid.uuid4().hex[:8]}"
    draft = NewReplyDraft(
        source_id=source_id,
        mode=ReplyMode.REPLY_ALL,
        body=f"<b>{marker}</b> & done",
        recipients=resolve_reply_recipients(
            source, ReplyMode.REPLY_ALL, writer.get_own_addresses()
        ),
    )
    draft_id = writer.create_reply(draft).id
    try:
        yield draft_id, marker, draft
    finally:
        client.delete(f"/me/messages/{quote(draft_id, safe='')}")


def fetch(client: GraphClient, message_id: str) -> Mapping[str, Any]:
    return client.get(f"/me/messages/{quote(message_id, safe='')}", {"$select": CHECK_FIELDS})


def addresses(message: Mapping[str, Any], field: str) -> list[str]:
    return [entry["emailAddress"]["address"].casefold() for entry in message.get(field, [])]


def test_comment_markup_is_saved_as_escaped_text(
    client: GraphClient, saved: tuple[str, str, NewReplyDraft]
) -> None:
    draft_id, marker, _ = saved

    body = fetch(client, draft_id)["body"]["content"]

    assert f"&lt;b&gt;{marker}&lt;/b&gt;" in body
    assert f"<b>{marker}</b>" not in body


def test_draft_is_in_the_original_conversation(
    client: GraphClient, source_id: str, saved: tuple[str, str, NewReplyDraft]
) -> None:
    draft_id, _, _ = saved

    draft = fetch(client, draft_id)

    assert draft["isDraft"] is True
    assert draft["conversationId"] == fetch(client, source_id)["conversationId"]


def test_draft_carries_exactly_the_recipients_sent(
    client: GraphClient, saved: tuple[str, str, NewReplyDraft]
) -> None:
    draft_id, _, sent = saved

    draft = fetch(client, draft_id)

    assert addresses(draft, "toRecipients") == [r.address.casefold() for r in sent.recipients.to]
    assert addresses(draft, "ccRecipients") == [r.address.casefold() for r in sent.recipients.cc]


def test_source_read_state_is_unchanged(client: GraphClient, source_id: str) -> None:
    writer = GraphReplyDraftWriter(client)
    was_read = fetch(client, source_id)["isRead"]
    source = writer.get_reply_source(source_id)
    draft = NewReplyDraft(
        source_id=source_id,
        mode=ReplyMode.REPLY,
        body="read-state check",
        recipients=resolve_reply_recipients(source, ReplyMode.REPLY, writer.get_own_addresses()),
    )
    draft_id = writer.create_reply(draft).id
    try:
        assert fetch(client, source_id)["isRead"] == was_read
    finally:
        client.delete(f"/me/messages/{quote(draft_id, safe='')}")
