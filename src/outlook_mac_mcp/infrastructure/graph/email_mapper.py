from collections.abc import Mapping
from datetime import datetime
from typing import Any

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError

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

    This is the only place where Graph's JSON is allowed to be loosely typed; every
    value is checked and converted here so no `Any` reaches the domain or the use cases.
    """
    return Email(
        id=_required_text(message, "id"),
        subject=_optional_text(message, "subject"),
        sender=_read_sender(message),
        received_at=_read_received_at(message),
        is_read=_required_flag(message, "isRead"),
        has_attachments=_required_flag(message, "hasAttachments"),
        preview=_optional_text(message, "bodyPreview"),
    )


def _read_sender(message: Mapping[str, Any]) -> EmailAddress:
    """Graph omits `from` on drafts and on some system messages, so an empty sender is valid."""
    sender = message.get("from")
    if not isinstance(sender, dict):
        return EmailAddress(address="")
    email_address = sender.get("emailAddress")
    if not isinstance(email_address, dict):
        return EmailAddress(address="")
    return EmailAddress(
        address=_optional_text(email_address, "address"),
        display_name=_optional_text(email_address, "name"),
    )


def _read_received_at(message: Mapping[str, Any]) -> datetime:
    raw = _required_text(message, "receivedDateTime")
    try:
        received_at = datetime.fromisoformat(raw)
    except ValueError as error:
        raise GraphResponseError("receivedDateTime was not an ISO 8601 timestamp") from error
    if received_at.tzinfo is None:
        raise GraphResponseError("receivedDateTime carried no time zone")
    return received_at


def _required_text(message: Mapping[str, Any], field: str) -> str:
    value = message.get(field)
    if not isinstance(value, str) or not value:
        raise GraphResponseError(f"the message carried no {field}")
    return value


def _optional_text(message: Mapping[str, Any], field: str) -> str:
    value = message.get(field)
    return value if isinstance(value, str) else ""


def _required_flag(message: Mapping[str, Any], field: str) -> bool:
    value = message.get(field)
    if not isinstance(value, bool):
        raise GraphResponseError(f"the message carried no {field}")
    return value
