from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.infrastructure.graph.json_fields import optional_text


def to_email_address(recipient: object) -> EmailAddress:
    """Read a Graph recipient, the `{emailAddress: {name, address}}` shape used for a
    message's `from` and an event's `organizer`.

    Graph omits the sender on drafts and some system messages and the organizer on some
    events, so an absent or malformed recipient is an empty address, never an error.
    """
    if not isinstance(recipient, dict):
        return EmailAddress(address="")
    email_address = recipient.get("emailAddress")
    if not isinstance(email_address, dict):
        return EmailAddress(address="")
    return read_email_address(email_address)


def read_email_address(email_address: Mapping[str, Any]) -> EmailAddress:
    """Read a Graph emailAddress object directly: the `{name, address}` shape a contact's
    `emailAddresses` collection holds unwrapped, unlike a message or event recipient.
    """
    return EmailAddress(
        address=optional_text(email_address, "address"),
        display_name=optional_text(email_address, "name"),
    )
