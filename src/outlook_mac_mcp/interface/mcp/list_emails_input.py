from typing import Annotated

from pydantic import Field

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.domain.sort_order import SortOrder
from outlook_mac_mcp.interface.mcp.email_filters_input import EmailFiltersInput

Sort = Annotated[
    SortOrder,
    Field(description="newest puts the most recent email first; oldest puts the earliest first."),
]
ListLimit = Annotated[
    int,
    Field(ge=MIN_LIMIT, le=MAX_LIMIT, description="Maximum number of emails to return."),
]


class ListEmailsInput(EmailFiltersInput):
    """The listing's input contract: the shared filters plus an order and a page size."""

    sort: Sort = SortOrder.NEWEST
    limit: ListLimit = DEFAULT_LIMIT

    def to_request(self, folder_ids: tuple[str, ...]) -> ListEmailsRequest:
        return ListEmailsRequest(
            filters=self.to_filters(folder_ids), sort=self.sort, limit=self.limit
        )
