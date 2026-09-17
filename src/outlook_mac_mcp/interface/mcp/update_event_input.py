from typing import Annotated

from pydantic import AwareDatetime, BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.event_changes import MAX_EVENT_ID_LENGTH, EventChanges
from outlook_mac_mcp.domain.new_event import (
    MAX_ATTENDEES,
    MAX_BODY_LENGTH,
    MAX_SUBJECT_LENGTH,
    MIN_SUBJECT_LENGTH,
)
from outlook_mac_mcp.interface.mcp.preview_event_input import (
    EMAIL_PATTERN,
    MAX_ADDRESS_LENGTH,
    MAX_LOCATION_LENGTH,
)

EventId = Annotated[
    str,
    Field(
        min_length=1,
        max_length=MAX_EVENT_ID_LENGTH,
        description="Graph id of the event, as returned by list_todays_events or "
        "list_upcoming_events.",
    ),
]
Subject = Annotated[
    str,
    Field(
        min_length=MIN_SUBJECT_LENGTH,
        max_length=MAX_SUBJECT_LENGTH,
        description="New title. Omit to leave the subject unchanged.",
    ),
]
Start = Annotated[
    AwareDatetime,
    Field(
        description="New start as ISO 8601 with a UTC offset. Omit to leave it unchanged; "
        "if given alone, must stay before the current end."
    ),
]
End = Annotated[
    AwareDatetime,
    Field(
        description="New end as ISO 8601 with a UTC offset. Omit to leave it unchanged; "
        "if given alone, must stay after the current start."
    ),
]
Location = Annotated[
    str,
    Field(
        max_length=MAX_LOCATION_LENGTH,
        description="New free-text place. Empty clears it; omit to leave it unchanged.",
    ),
]
Body = Annotated[
    str,
    Field(
        max_length=MAX_BODY_LENGTH,
        description=(
            "New plain-text body, up to 32,768 characters. Empty clears it; omit to leave "
            "it unchanged. No HTML, for the same untrusted-content reason as create_event."
        ),
    ),
]
Attendee = Annotated[str, Field(pattern=EMAIL_PATTERN, max_length=MAX_ADDRESS_LENGTH)]
Attendees = Annotated[
    tuple[Attendee, ...],
    Field(
        max_length=MAX_ATTENDEES,
        description="New full attendee list, replacing the current one, up to 50. Empty "
        "clears it; omit to leave it unchanged.",
    ),
]


class UpdateEventInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become an EventChanges.

    Every field but event_id is optional: a field left out of the call is left as is by
    the update, never cleared, the same None-means-unchanged rule EventChanges itself
    documents.
    """

    model_config = ConfigDict(frozen=True, extra="forbid")

    event_id: EventId
    subject: Subject | None = None
    start: Start | None = None
    end: End | None = None
    location: Location | None = None
    body: Body | None = None
    attendees: Attendees | None = None

    def to_changes(self) -> EventChanges:
        return EventChanges(
            event_id=self.event_id,
            subject=self.subject,
            start=self.start,
            end=self.end,
            location=self.location,
            body=self.body,
            attendees=(
                tuple(EmailAddress(address=address) for address in self.attendees)
                if self.attendees is not None
                else None
            ),
        )
