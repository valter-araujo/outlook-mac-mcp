from outlook_mac_mcp.application.search_contacts_request import SearchContactsRequest
from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.domain.page import Page


class InMemoryContactRepository:
    """Fake ContactRepository; applies the port's name-prefix-or-email-exact rule."""

    def __init__(self) -> None:
        self._contacts: list[Contact] = []

    def add(self, contact: Contact) -> None:
        self._contacts.append(contact)

    def search(self, request: SearchContactsRequest) -> Page[Contact]:
        needle = request.term.casefold()
        matches = [contact for contact in self._contacts if _matches(contact, needle)]
        return Page(items=tuple(matches[: request.limit]), total=len(matches), total_is_exact=True)


def _matches(contact: Contact, needle: str) -> bool:
    if contact.display_name.casefold().startswith(needle):
        return True
    return any(address.address.casefold() == needle for address in contact.email_addresses)
