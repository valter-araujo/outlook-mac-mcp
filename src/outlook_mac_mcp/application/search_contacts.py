from outlook_mac_mcp.application.ports.contact_repository import ContactRepository
from outlook_mac_mcp.application.search_contacts_request import SearchContactsRequest
from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.domain.page import Page


class SearchContacts:
    def __init__(self, contact_repository: ContactRepository) -> None:
        self._contact_repository = contact_repository

    def execute(self, request: SearchContactsRequest) -> Page[Contact]:
        return self._contact_repository.search(request)
