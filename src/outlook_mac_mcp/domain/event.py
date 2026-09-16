from dataclasses import dataclass
from datetime import datetime

from outlook_mac_mcp.domain.email_address import EmailAddress


@dataclass(frozen=True, slots=True)
class Event:
    """One occurrence on the calendar.

    `start` and `end` are time zone-aware instants. An all-day event spans midnight to
    midnight in its own zone and is not treated differently here; `is_all_day` only tells
    a reader that the clock times are not meaningful.

    `body` defaults to empty because a listing never selects it, the same reason `Email`
    carries no body; it is only ever populated by reading back a just-created event.
    """

    id: str
    subject: str
    start: datetime
    end: datetime
    is_all_day: bool
    location: str
    organizer: EmailAddress
    body: str = ""
