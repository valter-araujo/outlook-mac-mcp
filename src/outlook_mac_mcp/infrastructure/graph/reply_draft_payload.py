from typing import Any

from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.infrastructure.graph.mail_draft_payload import recipient_payload


def to_reply_draft_payload(draft: NewReplyDraft) -> dict[str, Any]:
    """Shape a reply the way createReply and createReplyAll expect it.

    The new text goes in `comment`, which Graph places above the quoted original, so the
    original's HTML is never read, rebuilt or rewritten here. The recipients are sent
    explicitly so the saved draft carries the list the preview showed, not one Graph
    recomputed. `message.body` must stay absent: Graph answers 400 when both it and
    `comment` are given.
    """
    return {
        "comment": draft.body,
        "message": {
            "toRecipients": [recipient_payload(recipient) for recipient in draft.recipients.to],
            "ccRecipients": [recipient_payload(recipient) for recipient in draft.recipients.cc],
        },
    }
