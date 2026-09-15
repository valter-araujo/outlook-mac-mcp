from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.interface.mcp.email_view import EmailView


class EmailPageView(BaseModel):
    """JSON-serializable projection of a Page of emails.

    `returned` duplicates the length of `items` on purpose: a model reading the output
    should not have to count a list to compare it with `total`.
    """

    model_config = ConfigDict(frozen=True)

    items: list[EmailView] = Field(description="The emails on this page.")
    returned: int = Field(description="How many emails this page holds.")
    total: int = Field(
        description=(
            "How many emails match in all. When total_is_exact is false this is a lower "
            "bound: the true number is this or more."
        )
    )
    total_is_exact: bool = Field(
        description="Whether total was counted fully rather than stopped at a bound."
    )

    @classmethod
    def from_page(cls, page: Page[Email]) -> "EmailPageView":
        return cls(
            items=[EmailView.from_email(email) for email in page.items],
            returned=len(page.items),
            total=page.total,
            total_is_exact=page.total_is_exact,
        )
