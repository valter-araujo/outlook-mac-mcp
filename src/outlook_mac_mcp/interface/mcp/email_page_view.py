from pydantic import Field

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.interface.mcp.email_view import EmailView
from outlook_mac_mcp.interface.mcp.page_view import PageView

FOLDER_DESCRIPTION = (
    "Every folder this call was actually scoped to, as a comma-separated list: one "
    "well-known folder's name, or all five of them when folder was all. Always the "
    "real names, never the literal word all, so two all results stay comparable "
    "without reading the code to know what all covered at the time."
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
