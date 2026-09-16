from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.email_size import EmailSize
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import to_email_address
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import (
    optional_text,
    required_datetime,
    required_text,
)

# PR_MESSAGE_SIZE, the MAPI property Outlook itself uses for a message's size in bytes.
# Requested via $expand=singleValueExtendedProperties($filter=id eq '...').
MESSAGE_SIZE_PROPERTY_ID = "Integer 0x0E08"
EXTENDED_PROPERTIES_FIELD = "singleValueExtendedProperties"


def to_email_size(message: Mapping[str, Any]) -> EmailSize:
    """Turn one Graph message resource, expanded with the size property, into an EmailSize."""
    return EmailSize(
        id=required_text(message, "id"),
        subject=optional_text(message, "subject"),
        sender=to_email_address(message.get("from")),
        received_at=required_datetime(message, "receivedDateTime"),
        size_bytes=_read_size_bytes(message),
    )


def _read_size_bytes(message: Mapping[str, Any]) -> int:
    """Graph always sends an extended property's value as a string, even for an integer
    property, so the size still needs its own parse; a missing or unparsable value is
    surfaced rather than defaulting to zero, which would silently understate a size.
    """
    properties = message.get(EXTENDED_PROPERTIES_FIELD)
    if not isinstance(properties, list):
        raise GraphResponseError(f"the message carried no {EXTENDED_PROPERTIES_FIELD}")
    for extended_property in properties:
        if not isinstance(extended_property, dict):
            continue
        if extended_property.get("id") != MESSAGE_SIZE_PROPERTY_ID:
            continue
        value = extended_property.get("value")
        if isinstance(value, str):
            try:
                return int(value)
            except ValueError as error:
                raise GraphResponseError(
                    f"the {MESSAGE_SIZE_PROPERTY_ID} value {value!r} was not an integer"
                ) from error
    raise GraphResponseError(f"the message carried no {MESSAGE_SIZE_PROPERTY_ID} property")
