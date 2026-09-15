from datetime import datetime

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.event import Event


class EventView(BaseModel):
    """JSON-serializable projection of an Event.

    Field for field, with the organizer flattened like the sender of an email. Times
    serialize as ISO 8601 with their offset, so a reader never has to guess the zone.
    """

    model_config = ConfigDict(frozen=True)

    id: str
    subject: str
    start: datetime
    end: datetime
    is_all_day: bool = Field(
        description="When true the clock times are midnight to midnight and not meaningful."
    )
    location: str
    organizer_address: str
    organizer_name: str

    @classmethod
    def from_event(cls, event: Event) -> "EventView":
        return cls(
            id=event.id,
            subject=event.subject,
            start=event.start,
            end=event.end,
            is_all_day=event.is_all_day,
            location=event.location,
            organizer_address=event.organizer.address,
            organizer_name=event.organizer.display_name,
        )
