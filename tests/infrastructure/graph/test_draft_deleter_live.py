r"""Live check of where GraphDraftDeleter's DELETE sends a draft.

The README and GraphDraftDeleter promise that a deleted draft lands in Deleted Items and
stays recoverable. Graph documents that for work accounts; a run on 2026-11-19 against a
personal account found fewer test drafts in Deleted Items than were deleted, so this
checks the promise directly.

This WRITES to the mailbox: it saves one draft addressed to an example.com address
(never sent), deletes it through GraphDraftDeleter, and looks for its unique subject in
Deleted Items. If found, the draft is left there for you to empty.

Not yet measured.

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
