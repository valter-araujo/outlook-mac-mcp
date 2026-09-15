from collections.abc import Mapping
from datetime import datetime
from typing import Any

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import to_email_address
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import (
    optional_text,
    required_flag,
    required_text,
)

TEXT_CONTENT_TYPE = "text"
MESSAGE_FIELDS = (
    "id",
    "subject",
    "from",
    "receivedDateTime",
    "isRead",
    "hasAttachments",
    "bodyPreview",
)


def to_email(message: Mapping[str, Any]) -> Email:
    """Turn one Graph message resource into a domain Email.

    Every value is checked and converted through the field readers, so no `Any` reaches
    the domain or the use cases.
    """
    return Email(
        id=required_text(message, "id"),
        subject=optional_text(message, "subject"),
        sender=to_email_address(message.get("from")),
        received_at=_read_received_at(message),
        is_read=required_flag(message, "isRead"),
        has_attachments=required_flag(message, "hasAttachments"),
        preview=optional_text(message, "bodyPreview"),
    )


def to_email_detail(message: Mapping[str, Any]) -> EmailDetail:
    """Turn one Graph message resource into an EmailDetail, body included."""
    return EmailDetail(email=to_email(message), body=_read_body(message))


def _read_body(message: Mapping[str, Any]) -> str:
    """Refuse a body that is not plain text.

    The Prefer header asks Graph for text. If it answers with HTML anyway, returning it
    as if it were text would hand markup to a model that was told it is reading text, so
    the mismatch is surfaced instead of hidden. A message with no body at all is normal.
    """
    body = message.get("body")
    if not isinstance(body, dict):
        return ""
    content_type = optional_text(body, "contentType")
    if content_type and content_type != TEXT_CONTENT_TYPE:
        raise GraphResponseError(f"Graph returned a {content_type} body, not plain text")
    return optional_text(body, "content")


def _read_received_at(message: Mapping[str, Any]) -> datetime:
    raw = required_text(message, "receivedDateTime")
    try:
        received_at = datetime.fromisoformat(raw)
    except ValueError as error:
        raise GraphResponseError("receivedDateTime was not an ISO 8601 timestamp") from error
    if received_at.tzinfo is None:
        raise GraphResponseError("receivedDateTime carried no time zone")
    return received_at
