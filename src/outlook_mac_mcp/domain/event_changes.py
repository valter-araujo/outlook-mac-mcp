from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.sensitivity import Sensitivity
from outlook_mac_mcp.domain.show_as import ShowAs

# Graph event ids are opaque and their length varies, so the bound exists to keep
# unbounded input out of a URL path, not to describe a format; mirrors MAX_EMAIL_ID_LENGTH.
MAX_EVENT_ID_LENGTH = 1024


@dataclass(frozen=True, slots=True)
class EventChanges:
    """What the caller wants to change about one existing event; a field left as None
    means "leave it as is", not "clear it".

    Per-field rules (subject length, body length, attendee count, start before end when
    only one boundary changes) are not enforced here: they are enforced once, on the
    merge of these changes into the current event, by constructing the resulting
    NewEvent and letting its own validation run. That is what "matches NewEvent's rules"
    means in practice — the same code decides, not a second copy of the same checks.
    What this validates is self-contained: the id, that an aware datetime was given
    where a datetime was given at all, and that something was actually asked to change.
    """

    event_id: str
    subject: str | None = None
    start: datetime | None = None
    end: datetime | None = None
    location: str | None = None
    body: str | None = None
    attendees: tuple[EmailAddress, ...] | None = None
    reminder_minutes_before_start: int | None = None
    sensitivity: Sensitivity | None = None
    show_as: ShowAs | None = None

    def __post_init__(self) -> None:
        if not self.event_id:
            raise InvalidRequestError("event_id must not be empty")
        if len(self.event_id) > MAX_EVENT_ID_LENGTH:
            raise InvalidRequestError(f"event_id must be at most {MAX_EVENT_ID_LENGTH} characters")
        if self.start is not None and self.start.tzinfo is None:
            raise InvalidRequestError("start must carry a time zone")
        if self.end is not None and self.end.tzinfo is None:
            raise InvalidRequestError("end must carry a time zone")
        if self.start is not None and self.end is not None and not self.start < self.end:
            raise InvalidRequestError("start must be before end")
        if not self.has_any_change():
            raise InvalidRequestError("at least one field must be supplied to update")

    def has_any_change(self) -> bool:
        return any(
            field is not None
            for field in (
                self.subject,
                self.start,
                self.end,
                self.location,
                self.body,
                self.attendees,
                self.reminder_minutes_before_start,
                self.sensitivity,
                self.show_as,
            )
        )
