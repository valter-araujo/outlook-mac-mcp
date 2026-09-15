from dataclasses import dataclass, field

from outlook_mac_mcp.application.limits import DEFAULT_LIMIT, ensure_limit_within_bounds
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.sort_order import SortOrder


@dataclass(frozen=True, slots=True)
class ListEmailsRequest:
    """A validated listing, and the single argument the repository port takes."""

    filters: EmailFilters = field(default_factory=EmailFilters)
    sort: SortOrder = SortOrder.NEWEST
    limit: int = DEFAULT_LIMIT

    def __post_init__(self) -> None:
        ensure_limit_within_bounds(self.limit)
