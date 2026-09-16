from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.interface.mcp.contact_view import ContactView
from outlook_mac_mcp.interface.mcp.page_view import PageView


class ContactPageView(PageView[ContactView]):
    @classmethod
    def from_page(cls, page: Page[Contact]) -> "ContactPageView":
        return cls.projected(page, ContactView.from_contact)
