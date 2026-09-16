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
# typeless form, and requesting it is unaffected by the finding below: both type
# keywords are sent, since either was observed to resolve the same property.
MESSAGE_SIZE_PROPERTY_ID_INTEGER = "Integer 0x0E08"
MESSAGE_SIZE_PROPERTY_ID_LONG = "Long 0x0E08"
EXTENDED_PROPERTIES_FIELD = "singleValueExtendedProperties"

# Which type keywords Graph accepts for this property when reading its id back. Matching
# must not require a specific one: see the note on _is_message_size_property below.
MESSAGE_SIZE_PROPERTY_TYPES = frozenset({"integer", "long"})
MESSAGE_SIZE_PROPERTY_TAG = 0x0E08


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
        if not _is_message_size_property(extended_property.get("id")):
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


def _is_message_size_property(property_id: object) -> bool:
    r"""Whether `property_id` names PR_MESSAGE_SIZE, regardless of type keyword or of how
    Graph formats the hex tag on the way back.

    Live check, 2026-09-16, personal Outlook.com mailbox, one real message: requesting
    $expand=singleValueExtendedProperties($filter=id eq 'Long 0x0E08') returned the
    property with `"id": "Long 0xe08"`; requesting 'Integer 0x0E08' on the same message
    returned `"id": "Integer 0xe08"` and the same value. Two things follow, neither of
    which an earlier, narrower live check had reason to catch:

    1. Graph does not echo an extended property's id back in the case or padding it was
       requested in — lowercase, no leading zero, regardless of what was sent. A mapper
       that compares the returned id against a fixed-case, zero-padded string constant
       (the bug this function replaces) never matches anything Graph actually sends,
       which is what turned every scanned message into a false "no size" skip.
    2. Either type keyword resolves the same property to the same value, so the type
       name is checked only against the small set Graph is known to use for it, never
       required to be one specific keyword.

    Do not go back to comparing `property_id` against a fixed string for this reason.
    """
    if not isinstance(property_id, str):
        return False
    type_name, _, tag = property_id.partition(" ")
    if type_name.casefold() not in MESSAGE_SIZE_PROPERTY_TYPES:
        return False
    try:
        return int(tag, 16) == MESSAGE_SIZE_PROPERTY_TAG
    except ValueError:
        return False
