from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.reply_source import ReplySource
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import (
    to_email_address,
    to_email_addresses,
)
from outlook_mac_mcp.infrastructure.graph.json_fields import optional_text, required_text

REPLY_SOURCE_FIELDS = ("id", "subject", "from", "replyTo", "toRecipients", "ccRecipients")


def to_reply_source(message: Mapping[str, Any]) -> ReplySource:
    return ReplySource(
        id=required_text(message, "id"),
        subject=optional_text(message, "subject"),
        sender=to_email_address(message.get("from")),
        reply_to=to_email_addresses(message.get("replyTo")),
        to=to_email_addresses(message.get("toRecipients")),
        cc=to_email_addresses(message.get("ccRecipients")),
    )
