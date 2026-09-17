from typing import Annotated

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.application.preview_event_deletion import EventDeletionRequest
from outlook_mac_mcp.domain.event_changes import MAX_EVENT_ID_LENGTH

EventId = Annotated[
    str,
    Field(
        min_length=1,
        max_length=MAX_EVENT_ID_LENGTH,
        description="Graph id of the event, as returned by list_todays_events or "
        "list_upcoming_events.",
    ),
]


class EventDeletionInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a request."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    event_id: EventId

    def to_request(self) -> EventDeletionRequest:
        return EventDeletionRequest(event_id=self.event_id)
