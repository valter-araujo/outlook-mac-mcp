from typing import Annotated

from pydantic import AwareDatetime, BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_event import (
    MAX_ATTENDEES,
    MAX_BODY_LENGTH,
    MAX_SUBJECT_LENGTH,
    MIN_SUBJECT_LENGTH,
    NewEvent,
)

# Deliberately loose: one @, no whitespace, a dot in the domain. Graph is the authority
# on deliverability; this only keeps obvious non-addresses out of the payload.
EMAIL_PATTERN = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
MAX_ADDRESS_LENGTH = 320
MAX_LOCATION_LENGTH = 255

Subject = Annotated[
    str,
    Field(
        min_length=MIN_SUBJECT_LENGTH,
        max_length=MAX_SUBJECT_LENGTH,
        description="Title of the event.",
    ),
]
Start = Annotated[
    AwareDatetime,
    Field(description="Start as ISO 8601 with a UTC offset, e.g. 2026-09-15T09:00:00-03:00."),
]
End = Annotated[
    AwareDatetime,
    Field(description="End as ISO 8601 with a UTC offset; must be after start."),
]
IsAllDay = Annotated[
    bool,
    Field(description="When true, start and end must be midnights and end is exclusive."),
]
Location = Annotated[
    str, Field(max_length=MAX_LOCATION_LENGTH, description="Free-text place, or empty.")
]
Body = Annotated[
    str,
    Field(
        max_length=MAX_BODY_LENGTH,
        description=(
            "Plain text only, up to 32,768 characters, or empty. No HTML: there is no "
            "content-type choice, since a body can originate from untrusted content "
            "(an email or a web page) and markup there is the same injection surface "
            "get_email's body warning covers."
        ),
    ),
]
Attendee = Annotated[str, Field(pattern=EMAIL_PATTERN, max_length=MAX_ADDRESS_LENGTH)]
# A tuple so the tool signature can carry an immutable empty default.
Attendees = Annotated[
    tuple[Attendee, ...],
    Field(max_length=MAX_ATTENDEES, description="Email addresses to invite, up to 50."),
]


class PreviewEventInput(BaseModel):
    """The tool's input contract, and the only place MCP arguments become a NewEvent."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    subject: Subject
    start: Start
    end: End
    is_all_day: IsAllDay = False
    location: Location = ""
    body: Body = ""
    attendees: Attendees = ()

    def to_new_event(self) -> NewEvent:
        return NewEvent(
            subject=self.subject,
            start=self.start,
            end=self.end,
            is_all_day=self.is_all_day,
            location=self.location,
            body=self.body,
            attendees=tuple(EmailAddress(address=address) for address in self.attendees),
        )
