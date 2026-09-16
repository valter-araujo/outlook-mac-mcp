from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.infrastructure.graph.email_address_mapper import read_email_address
from outlook_mac_mcp.infrastructure.graph.json_fields import optional_text, required_text

CONTACT_FIELDS = ("id", "displayName", "emailAddresses")


def to_contact(contact: Mapping[str, Any]) -> Contact:
    return Contact(
        id=required_text(contact, "id"),
        display_name=optional_text(contact, "displayName"),
        email_addresses=tuple(_read_email_addresses(contact)),
    )


def _read_email_addresses(contact: Mapping[str, Any]) -> list[Any]:
    addresses = contact.get("emailAddresses")
    if not isinstance(addresses, list):
        return []
    return [read_email_address(item) for item in addresses if isinstance(item, dict)]
