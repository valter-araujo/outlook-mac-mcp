from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.interface.mcp.email_view import EmailView
from outlook_mac_mcp.interface.mcp.page_view import PageView


class EmailPageView(PageView[EmailView]):
    @classmethod
    def from_page(cls, page: Page[Email]) -> "EmailPageView":
        return cls.projected(page, EmailView.from_email)
