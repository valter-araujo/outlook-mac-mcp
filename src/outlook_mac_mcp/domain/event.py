from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.sensitivity import Sensitivity
from outlook_mac_mcp.domain.show_as import ShowAs


@dataclass(frozen=True, slots=True)
class Event:
    """One occurrence on the calendar.

    `start` and `end` are time zone-aware instants. An all-day event spans midnight to
    midnight in its own zone and is not treated differently here; `is_all_day` only tells
    a reader that the clock times are not meaningful.

    `body` and `attendees` default to empty because a listing never selects either, the
    same reason `Email` carries no body; both are only ever populated by a read that
    asks for them specifically, such as fetching one event by id.

    `reminder_minutes_before_start`, `sensitivity` and `show_as` default to None for the
    same reason: a listing never selects them, so None means "not read here", not "not
    set on the calendar".
    """

    id: str
    subject: str
    start: datetime
    end: datetime
    is_all_day: bool
    location: str
    organizer: EmailAddress
    body: str = ""
    attendees: tuple[EmailAddress, ...] = ()
    reminder_minutes_before_start: int | None = None
    sensitivity: Sensitivity | None = None
    show_as: ShowAs | None = None
