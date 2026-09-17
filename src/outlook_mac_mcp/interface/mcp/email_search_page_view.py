from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email_search_page import EmailSearchPage
from outlook_mac_mcp.interface.mcp.email_page_view import FOLDER_DESCRIPTION
from outlook_mac_mcp.interface.mcp.email_view import EmailView


class EmailSearchPageView(BaseModel):
    """JSON-serializable projection of an EmailSearchPage.

    search_emails' own response shape, not PageView: it is the only listing tool with a
    next_page_token, and PageView is shared by every other one.
    """

    model_config = ConfigDict(frozen=True)

    items: list[EmailView] = Field(description="The items on this page.")
    returned: int = Field(description="How many items this page holds.")
    total: int = Field(
        description=(
            "How many items match in all. When total_is_exact is false this is a lower "
            "bound: the true number is this or more."
        )
    )
    total_is_exact: bool = Field(
        description="Whether total was counted fully rather than stopped at a bound."
    )
    next_page_token: str | None = Field(
        description=(
            "Pass this to search_emails' page_token to get the next page of this same "
            "search. Absent when this is the last page, or when folder was all -- a "
            "multi-folder search cannot be continued, only re-run."
        )
    )
    folder: str = Field(description=FOLDER_DESCRIPTION)

    @classmethod
    def from_page(cls, page: EmailSearchPage, folder: str) -> "EmailSearchPageView":
        return cls(
            items=[EmailView.from_email(email) for email in page.items],
            returned=len(page.items),
            total=page.total,
            total_is_exact=page.total_is_exact,
            next_page_token=page.next_page_token,
            folder=folder,
        )
