from typing import Any

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft


def to_mail_draft_payload(new_draft: NewMailDraft) -> dict[str, Any]:
    """Shape a NewMailDraft the way POST /me/messages expects it.

    The body is always declared as text. A message created through this endpoint stays a
    draft until someone sends it; nothing in the payload can change that.
    """
    return {
        "subject": new_draft.subject,
        "body": {"contentType": "text", "content": new_draft.body},
        "toRecipients": [_recipient(recipient) for recipient in new_draft.to],
        "ccRecipients": [_recipient(recipient) for recipient in new_draft.cc],
    }


def _recipient(recipient: EmailAddress) -> dict[str, Any]:
    return {"emailAddress": {"address": recipient.address}}
