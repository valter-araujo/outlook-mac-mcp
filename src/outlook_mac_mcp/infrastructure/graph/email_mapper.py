from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import to_email_address
from outlook_mac_mcp.infrastructure.graph.json_fields import (
    optional_text,
    optional_text_body,
    required_datetime,
    required_flag,
    required_text,
)

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
        received_at=required_datetime(message, "receivedDateTime"),
        is_read=required_flag(message, "isRead"),
        has_attachments=required_flag(message, "hasAttachments"),
        preview=optional_text(message, "bodyPreview"),
    )


def to_email_detail(message: Mapping[str, Any]) -> EmailDetail:
    """Turn one Graph message resource into an EmailDetail, body included.

    The Prefer header asks Graph for the body as text; optional_text_body is what
    refuses an HTML answer instead of quietly handing markup to a reader expecting text.
    """
    return EmailDetail(email=to_email(message), body=optional_text_body(message))
