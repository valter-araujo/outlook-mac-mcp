from typing import Protocol

from outlook_mac_mcp.application.search_contacts_request import SearchContactsRequest
from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.domain.page import Page


class ContactRepository(Protocol):
    def search(self, request: SearchContactsRequest) -> Page[Contact]:
        """Return the contacts matching `request`.

        A contact matches when its display name starts with the term, or when the term
        is exactly one of its email addresses: Microsoft Graph does not support a prefix
        match on email address, only an exact one, so a term that is not a complete
        address matches by name alone.
        """
        ...
