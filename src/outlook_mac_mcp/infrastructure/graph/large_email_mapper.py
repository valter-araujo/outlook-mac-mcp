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
# Its documented proptag id form is "{type} {proptag}"; Microsoft's own docs list no
# typeless form. Which type name a tenant actually uses for this property is not
# something the docs settle, so both are requested and both are accepted here.
MESSAGE_SIZE_PROPERTY_ID_INTEGER = "Integer 0x0E08"
MESSAGE_SIZE_PROPERTY_ID_LONG = "Long 0x0E08"
MESSAGE_SIZE_PROPERTY_IDS = (MESSAGE_SIZE_PROPERTY_ID_INTEGER, MESSAGE_SIZE_PROPERTY_ID_LONG)
EXTENDED_PROPERTIES_FIELD = "singleValueExtendedProperties"


def to_email_size(message: Mapping[str, Any]) -> EmailSize | None:
    """Turn one Graph message resource, expanded with the size property, into an
    EmailSize, or None when the message carries neither form of the property.

    Absence is tolerated on purpose: a message can legitimately carry neither id, and
    the caller counts that as skipped rather than treating it as a broken response. A
    property that is present but carries a value that cannot be read as an integer is a
    different failure, one this still raises on, because that is not absence, it is
    Graph answering with something this project has no way to interpret as a size.
    """
    size_bytes = _read_size_bytes(message)
    if size_bytes is None:
        return None
    return EmailSize(
        id=required_text(message, "id"),
        subject=optional_text(message, "subject"),
        sender=to_email_address(message.get("from")),
        received_at=required_datetime(message, "receivedDateTime"),
        size_bytes=size_bytes,
    )


def _read_size_bytes(message: Mapping[str, Any]) -> int | None:
    """Graph always sends an extended property's value as a string, even for an integer
    property, so the size still needs its own parse.
    """
    properties = message.get(EXTENDED_PROPERTIES_FIELD)
    if not isinstance(properties, list):
        return None
    for extended_property in properties:
        if not isinstance(extended_property, dict):
            continue
        if extended_property.get("id") not in MESSAGE_SIZE_PROPERTY_IDS:
            continue
        value = extended_property.get("value")
        if not isinstance(value, str):
            raise GraphResponseError(
                f"a message size property carried a non-string value: {value!r}"
            )
        try:
            return int(value)
        except ValueError as error:
            raise GraphResponseError(
                f"a message size property value {value!r} was not an integer"
            ) from error
    return None
