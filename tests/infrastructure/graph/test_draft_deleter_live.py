r"""Live check that a draft deleted through GraphDraftDeleter lands in Deleted Items.

The README and GraphDraftDeleter promise that a deleted draft stays visible and
recoverable in Deleted Items, so this checks the promise directly.

This WRITES to the mailbox: it saves one draft addressed to an example.com address
(never sent), deletes it through GraphDraftDeleter, and looks for its unique subject in
Deleted Items. If found, the draft is left there for you to empty.

Measured 2026-11-19 against a personal Outlook.com mailbox, while the deleter still sent
an HTTP DELETE: FAILED. The draft was in neither Deleted Items, Drafts nor a
mailbox-wide subject search; listing Recoverable Items returned 503, so where it went
is unconfirmed. That led to the deleter moving the draft to `deleteditems` instead.
Re-run 2026-11-21 after the move change, same mailbox: PASSED, the draft was found in
Deleted Items.

Run with:
    OUTLOOK_MCP_ENABLE_MAIL_DRAFT=true \
    uv run pytest -m integration tests/infrastructure/graph/test_draft_deleter_live.py
"""

import os
import uuid

import pytest

from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.draft_deleter import GraphDraftDeleter
from outlook_mac_mcp.infrastructure.settings import load_settings

pytestmark = pytest.mark.integration

DELETED_ITEMS = "/me/mailFolders/deleteditems/messages"
RECENT_PAGE = 50


def test_a_deleted_draft_lands_in_deleted_items() -> None:
    client = GraphClient(DeviceCodeAuthenticator.from_settings(load_settings(os.environ)))
    subject = f"live-check-{uuid.uuid4().hex[:8]}"
    created = client.post(
        "/me/messages",
        {"subject": subject, "toRecipients": [{"emailAddress": {"address": "a@example.com"}}]},
    )
    deleter = GraphDraftDeleter(client)
    deleter.delete(deleter.get_draft(str(created["id"])).id)

    recent = client.get(
        DELETED_ITEMS,
        {"$select": "subject", "$orderby": "lastModifiedDateTime desc", "$top": RECENT_PAGE},
    )

    subjects = [item.get("subject") for item in recent.get("value", [])]
    assert subject in subjects
