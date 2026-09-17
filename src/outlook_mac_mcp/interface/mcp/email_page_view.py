from pydantic import Field

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.interface.mcp.email_view import EmailView
from outlook_mac_mcp.interface.mcp.page_view import PageView

FOLDER_DESCRIPTION = (
    "The folder value this call was scoped to: one well-known folder, or all when every "
    "well-known folder was searched and merged."
)


class EmailPageView(PageView[EmailView]):
    folder: str = Field(description=FOLDER_DESCRIPTION)

    @classmethod
    def from_page(cls, page: Page[Email], folder: str) -> "EmailPageView":
        return cls(
            items=[EmailView.from_email(email) for email in page.items],
            returned=len(page.items),
            total=page.total,
            total_is_exact=page.total_is_exact,
            folder=folder,
        )
