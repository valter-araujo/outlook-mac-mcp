from collections.abc import Callable
from typing import Self

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.page import Page


class PageView[ItemView](BaseModel):
    """JSON-serializable projection of a Page, shared by every listing tool.

    `returned` duplicates the length of `items` on purpose: a model reading the output
    should not have to count a list to compare it with `total`.
    """

    model_config = ConfigDict(frozen=True)

    items: list[ItemView] = Field(description="The items on this page.")
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

    @classmethod
    def projected[Item](cls, page: Page[Item], project: Callable[[Item], ItemView]) -> Self:
        return cls(
            items=[project(item) for item in page.items],
            returned=len(page.items),
            total=page.total,
            total_is_exact=page.total_is_exact,
        )
