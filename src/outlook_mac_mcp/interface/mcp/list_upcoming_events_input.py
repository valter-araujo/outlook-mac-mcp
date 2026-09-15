from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.list_upcoming_events import (
    DEFAULT_UPCOMING_DAYS,
    MAX_UPCOMING_DAYS,
    MIN_UPCOMING_DAYS,
    ListUpcomingEventsRequest,
)

Days = Annotated[
    int,
    Field(
        ge=MIN_UPCOMING_DAYS,
        le=MAX_UPCOMING_DAYS,
        description="How many days ahead to look, counted from now.",
    ),
]


class ListUpcomingEventsInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    days: Days = DEFAULT_UPCOMING_DAYS

    def to_request(self) -> ListUpcomingEventsRequest:
        return ListUpcomingEventsRequest(days=self.days)
